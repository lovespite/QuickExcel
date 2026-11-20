#!/usr/bin/env python3
"""Focused release validation and draft-publication failure tests; no GitHub writes."""
import copy
import importlib.util
import json
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch

spec = importlib.util.spec_from_file_location("github_release", Path(__file__).with_name("github-release.py"))
release = importlib.util.module_from_spec(spec)
spec.loader.exec_module(release)


class ReleaseTests(unittest.TestCase):
    def setUp(self):
        artifacts = release.ROOT / "artifacts"
        artifacts.mkdir(exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(prefix="github-release-test-", dir=artifacts)
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.root_patch = patch.object(release, "ROOT", self.root)
        self.root_patch.start()
        self.addCleanup(self.root_patch.stop)
        project = self.root / "QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj"
        project.parent.mkdir()
        project.write_text("<Project><PropertyGroup><Version>4.2.0</Version></PropertyGroup></Project>")
        self.sha = "a" * 40
        self.metadata = {"tag": "v4.2.0", "version": "4.2.0", "sha": self.sha}
        self.directory = self.root / "downloads"
        self.directory.mkdir()
        for name in ("QuickExcel-4.2.0-win-x64.zip", "QuickExcel-4.2.0-osx-arm64.tar.gz"):
            archive = self.directory / name
            archive.write_bytes((name + " verified binary").encode())
            (self.directory / (name + ".sha256")).write_text(release.sha256(archive) + "  " + name + "\n")
        self.assets = release.validate_assets(self.directory, "4.2.0")
        self.draft = {"draft": True, "assets": [dict(name=name, state="uploaded", size=asset["size"],
                                                     digest=asset["digest"]) for name, asset in self.assets.items()]}

    def test_tag_must_match_product_version_and_canonical_format(self):
        for tag in ("v4.1.0", "4.2.0", "v04.2.0", "v4.2.0\n", "v4.2.0-rc.1", "main"):
            with self.subTest(tag=tag), self.assertRaises(ValueError), patch.object(release, "run") as run:
                release.validate(tag)
            run.assert_not_called()

    def test_checkout_must_match_tag_and_original_workflow_commit(self):
        with patch.object(release, "run", side_effect=[self.sha, "b" * 40]), self.assertRaisesRegex(ValueError, "exact commit"):
            release.validate("v4.2.0")
        with patch.object(release, "run", return_value=self.sha), self.assertRaisesRegex(ValueError, "moved"):
            release.validate("v4.2.0", "b" * 40)
        with patch.object(release, "run", return_value=self.sha):
            self.assertEqual(self.metadata, release.validate("v4.2.0", self.sha))

    def test_missing_or_unexpected_platform_asset_is_rejected(self):
        extra = self.directory / "unverified.zip"
        extra.write_bytes(b"unexpected")
        with self.assertRaisesRegex(ValueError, "exactly"):
            release.validate_assets(self.directory, "4.2.0")
        extra.unlink()
        next(self.directory.glob("*.sha256")).unlink()
        with self.assertRaisesRegex(ValueError, "exactly"):
            release.validate_assets(self.directory, "4.2.0")

    def test_bad_hash_or_sidecar_filename_is_rejected(self):
        archive = next(self.directory.glob("*.zip"))
        archive.write_bytes(b"tampered")
        with self.assertRaisesRegex(ValueError, "Checksum"):
            release.validate_assets(self.directory, "4.2.0")
        archive.with_suffix(".zip.sha256").write_text(release.sha256(archive) + "  another-file.zip")
        with self.assertRaisesRegex(ValueError, "Checksum"):
            release.validate_assets(self.directory, "4.2.0")
        archive.with_suffix(".zip.sha256").write_text(release.sha256(archive) + "  " + archive.name.lower())
        with self.assertRaisesRegex(ValueError, "Checksum"):
            release.validate_assets(self.directory, "4.2.0")

    def publish_with_mocks(self, drafts, command=None, remote_sha=None):
        self.env = patch.dict("os.environ", {"GH_REPO": "owner/QuickExcel"})
        self.env.start()
        self.addCleanup(self.env.stop)
        validate = patch.object(release, "validate", return_value=self.metadata)
        validate.start()
        self.addCleanup(validate.stop)
        tag = patch.object(release, "remote_tag_sha", return_value=remote_sha or self.sha)
        tag.start()
        self.addCleanup(tag.stop)
        find = patch.object(release, "find_release", side_effect=drafts)
        find.start()
        self.addCleanup(find.stop)
        run = patch.object(release, "run", side_effect=command, return_value="")
        mocked = run.start()
        self.addCleanup(run.stop)
        return mocked

    def test_new_release_is_draft_until_all_verified_uploads_complete(self):
        run = self.publish_with_mocks([None, self.draft])
        release.publish("v4.2.0", self.directory)
        calls = [call.args for call in run.call_args_list]
        self.assertEqual(["create", "upload", "edit"], [args[2] for args in calls])
        self.assertIn("--draft", calls[0])
        self.assertIn("--verify-tag", calls[0])
        self.assertIn("--generate-notes", calls[0])
        self.assertIn("--draft=false", calls[-1])

    def test_failed_partial_upload_never_publishes(self):
        def fail_upload(*args):
            if args[2] == "upload":
                raise subprocess.CalledProcessError(1, args, stderr="network failure")
            return ""
        run = self.publish_with_mocks([None], fail_upload)
        with self.assertRaises(subprocess.CalledProcessError):
            release.publish("v4.2.0", self.directory)
        self.assertEqual(["create", "upload"], [call.args[2] for call in run.call_args_list])

    def test_existing_draft_can_resume_but_public_release_cannot_change(self):
        run = self.publish_with_mocks([self.draft, self.draft])
        release.publish("v4.2.0", self.directory)
        self.assertEqual(["upload", "edit"], [call.args[2] for call in run.call_args_list])
        run.reset_mock()
        with patch.object(release, "find_release", return_value={"draft": False}), self.assertRaisesRegex(ValueError, "already public"):
            release.publish("v4.2.0", self.directory)
        run.assert_not_called()

    def test_moved_remote_tag_prevents_release_mutations(self):
        run = self.publish_with_mocks([], remote_sha="b" * 40)
        with self.assertRaisesRegex(ValueError, "no longer points"):
            release.publish("v4.2.0", self.directory)
        run.assert_not_called()

    def test_missing_or_wrong_remote_digest_never_publishes(self):
        for bad_digest in (None, "sha256:" + "0" * 64):
            draft = copy.deepcopy(self.draft)
            draft["assets"][0]["digest"] = bad_digest
            with self.subTest(digest=bad_digest), self.assertRaisesRegex(ValueError, "upload verification"):
                release.verify_uploaded(draft, self.assets)
        draft = copy.deepcopy(self.draft)
        draft["assets"][0]["digest"] = None
        run = self.publish_with_mocks([self.draft, draft])
        with self.assertRaisesRegex(ValueError, "upload verification"):
            release.publish("v4.2.0", self.directory)
        self.assertEqual(["upload"], [call.args[2] for call in run.call_args_list])

    def test_api_lookup_failure_is_not_treated_as_absent_release(self):
        with patch.object(release, "run", side_effect=subprocess.CalledProcessError(1, ["gh"])), self.assertRaises(subprocess.CalledProcessError):
            release.find_release("owner/QuickExcel", "v4.2.0")

    def test_annotated_remote_tag_is_resolved_to_commit(self):
        with patch.object(release, "run", side_effect=[json.dumps({"object": {"type": "tag", "sha": "b" * 40}}),
                                                       json.dumps({"object": {"type": "commit", "sha": self.sha}})]):
            self.assertEqual(self.sha, release.remote_tag_sha("owner/QuickExcel", "v4.2.0"))


if __name__ == "__main__":
    unittest.main()
