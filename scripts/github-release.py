#!/usr/bin/env python3
"""Validate a release tag and publish the two verified platform archives via gh."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
TAG = re.compile(r"v(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\Z")


def run(*arguments):
    return subprocess.run(arguments, cwd=ROOT, check=True, capture_output=True,
                          text=True, encoding="utf-8").stdout.strip()


def validate(tag, expected_sha=None):
    if not TAG.fullmatch(tag):
        raise ValueError("Release tags must have the form vMAJOR.MINOR.PATCH, for example v4.2.0.")
    project = ROOT / "QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj"
    version = ET.parse(project).findtext("PropertyGroup/Version")
    if tag != "v" + str(version):
        raise ValueError(f"Tag {tag} does not match product version {version}.")
    sha = run("git", "rev-parse", "--verify", f"refs/tags/{tag}^{{commit}}")
    if sha != run("git", "rev-parse", "HEAD"):
        raise ValueError("The checkout must be the exact commit identified by the release tag.")
    if expected_sha and sha != expected_sha:
        raise ValueError("The release tag moved after validation; refusing to publish another commit.")
    return {"tag": tag, "version": version, "sha": sha}


def sha256(path):
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def validate_assets(directory, version):
    names = [f"QuickExcel-{version}-win-x64.zip", f"QuickExcel-{version}-osx-arm64.tar.gz"]
    expected = {name for archive in names for name in (archive, archive + ".sha256")}
    if not directory.is_dir() or {p.name for p in directory.iterdir()} != expected:
        raise ValueError("Release directory must contain exactly both platform archives and their .sha256 files.")
    assets = {}
    for name in sorted(expected):
        path = directory / name
        if path.is_symlink() or not path.is_file() or path.stat().st_size == 0:
            raise ValueError(f"Invalid release asset: {name}")
        assets[name] = {"path": path, "size": path.stat().st_size, "digest": "sha256:" + sha256(path)}
    for name in names:
        checksum = (directory / (name + ".sha256")).read_text(encoding="utf-8-sig").strip()
        record = re.fullmatch(r"([A-Fa-f0-9]{64})  (.+)", checksum)
        if not record or record[1].lower() != assets[name]["digest"][7:] or record[2] != name:
            raise ValueError(f"Checksum or checksum filename does not match {name}.")
    return assets


def remote_tag_sha(repo, tag):
    entry = json.loads(run("gh", "api", f"repos/{repo}/git/ref/tags/{tag}"))["object"]
    for _ in range(16):
        if entry["type"] == "commit":
            return entry["sha"]
        if entry["type"] != "tag":
            break
        entry = json.loads(run("gh", "api", f"repos/{repo}/git/tags/{entry['sha']}"))["object"]
    raise ValueError("GitHub release tag does not resolve to a commit.")


def find_release(repo, tag):
    # List includes drafts and distinguishes an absent release from API/auth failures.
    result = run("gh", "api", f"repos/{repo}/releases", "--paginate", "--jq",
                 f'.[] | select(.tag_name == "{tag}")')
    return json.loads(result) if result else None


def verify_uploaded(release, assets):
    remote = {asset["name"]: asset for asset in release["assets"]}
    if set(remote) != set(assets):
        raise ValueError("The draft does not contain exactly the expected release assets.")
    for name, local in assets.items():
        uploaded = remote[name]
        if (uploaded.get("state") != "uploaded" or uploaded.get("size") != local["size"]
                or uploaded.get("digest") != local["digest"]):
            raise ValueError(f"GitHub upload verification failed for {name}; release remains a draft.")


def publish(tag, directory, expected_sha=None):
    metadata = validate(tag, expected_sha)
    assets = validate_assets(directory, metadata["version"])
    repo = os.environ.get("GH_REPO") or os.environ.get("GITHUB_REPOSITORY", "")
    if not re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repo):
        raise ValueError("Set GH_REPO or GITHUB_REPOSITORY to owner/repository.")
    if remote_tag_sha(repo, tag) != metadata["sha"]:
        raise ValueError("GitHub tag no longer points to the validated source commit.")
    release = find_release(repo, tag)
    if release and not release["draft"]:
        raise ValueError(f"Release {tag} is already public. Its downloads will not be replaced; use a new version tag.")
    if release and {asset["name"] for asset in release["assets"]} - assets.keys():
        raise ValueError("Existing draft has unrelated assets; inspect it before rerunning publication.")
    if not release:
        notes = ROOT / "artifacts" / f"github-release-notes-{tag}.md"
        notes.parent.mkdir(parents=True, exist_ok=True)
        notes.write_text(
            "Windows x64: extract the ZIP and run `application/quickxl.exe`.\n\n"
            "macOS Apple Silicon: extract the tar.gz and run `QuickExcel/Start-QuickExcel.sh`. "
            "This is a raw self-contained package, not a Developer ID signed or notarized .app. "
            "Persistent script trust requires a signed bundle; this package supports explicit one-time consent.\n\n"
            "Both packages passed their production self-check. SHA-256 sidecars verify the downloads. "
            "Physical input, printing and clean-device acceptance remain separate.\n\n",
            encoding="utf-8")
        run("gh", "release", "create", tag, "--repo", repo, "--draft", "--verify-tag",
            "--title", f"QuickExcel {metadata['version']}", "--generate-notes", "--notes-file", str(notes))
    # Replacing assets is allowed only while the release is a draft. A partial
    # upload leaves that draft available for a safe workflow retry.
    run("gh", "release", "upload", tag, "--repo", repo, "--clobber",
        *(str(asset["path"]) for asset in assets.values()))
    release = find_release(repo, tag)
    if not release or not release["draft"]:
        raise ValueError("Release draft state changed while uploading; publication stopped.")
    verify_uploaded(release, assets)
    if remote_tag_sha(repo, tag) != metadata["sha"]:
        raise ValueError("GitHub tag moved during upload; release remains a draft.")
    run("gh", "release", "edit", tag, "--repo", repo, "--draft=false")
    print(f"Published {tag}: https://github.com/{repo}/releases/tag/{tag}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    for command in ("validate", "publish"):
        subparser = subparsers.add_parser(command)
        subparser.add_argument("tag")
        subparser.add_argument("--expected-sha")
        if command == "publish":
            subparser.add_argument("--directory", type=Path, default=ROOT / "artifacts/github-release")
    args = parser.parse_args()
    if args.command == "publish":
        publish(args.tag, args.directory, args.expected_sha)
    else:
        metadata = validate(args.tag, args.expected_sha)
        output = os.environ.get("GITHUB_OUTPUT")
        if output:
            with open(output, "a", encoding="utf-8") as target:
                for name, value in metadata.items():
                    target.write(f"{name}={value}\n")
        print(json.dumps(metadata))


if __name__ == "__main__":
    try:
        main()
    except (ValueError, OSError, ET.ParseError, subprocess.CalledProcessError) as error:
        print(f"Release failed: {error}", file=sys.stderr)
        if isinstance(error, subprocess.CalledProcessError) and error.stderr:
            print(error.stderr, file=sys.stderr)
        sys.exit(1)
