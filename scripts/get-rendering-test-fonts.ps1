[CmdletBinding()]
param([string]$OutputDirectory = 'artifacts/skia-p0/fonts')
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$destination = [IO.Path]::GetFullPath((Join-Path $root $OutputDirectory))
if (-not $destination.StartsWith((Join-Path $root 'artifacts') + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Fonts must stay under repository artifacts.' }
New-Item -ItemType Directory -Force -Path $destination | Out-Null
$commit = 'ffebf8c1ee449e544955a7e813c54f9b73848eac'
$fonts = @(
    @('NotoSans-Regular.ttf', 'NotoSans', 'B85C38ECEA8A7CFB39C24E395A4007474FA5A4FC864F6EE33309EB4948D232D5'),
    @('NotoSans-Bold.ttf', 'NotoSans', 'C976E4B1B99EDC88775377FCC21692CA4BFA46B6D6CA6522BFDA505B28FF9D6A'),
    @('NotoSans-Italic.ttf', 'NotoSans', '36CFF144DF01309DAB648BEA71BAFF9BB074026914AFE63AEACC8BC90B67A28B'),
    @('NotoMusic-Regular.ttf', 'NotoMusic', '5982F8C6A10C50809F879A28911A214A066197871EE8CAA544A86121D0452FD5'),
    @('NotoSansArabic-Regular.ttf', 'NotoSansArabic', 'CEEA25B464A656DC3B26849BAB9356740401AF62AEDF1BFA8B7F0D9B75925B1B'),
    @('NotoSansHebrew-Regular.ttf', 'NotoSansHebrew', 'A7FA16FFFB27BEDB060A0866267C29E9859AEB9C21CC33F5B3AAF6EB062ECA85'),
    @('NotoSansSymbols2-Regular.ttf', 'NotoSansSymbols2', '630846D528DBE4C4981370A4D0A9475A1FD1491A129BB411F8E157CDB5DE13C6'),
    @('NotoSansCJKsc-Regular.otf', 'CJK', '2C76254F6FC379FDDFCE0A7E84FB5385BB135D3E399294F6EEB6680D0365B74B')
)
$manifest = foreach ($font in $fonts) {
    $url = if ($font[1] -eq 'CJK') { 'https://raw.githubusercontent.com/notofonts/noto-cjk/Sans2.004/Sans/OTF/SimplifiedChinese/NotoSansCJKsc-Regular.otf' }
        else { "https://raw.githubusercontent.com/notofonts/noto-fonts/$commit/hinted/ttf/$($font[1])/$($font[0])" }
    $path = Join-Path $destination $font[0]
    if (-not (Test-Path -LiteralPath $path)) { Invoke-WebRequest $url -OutFile $path }
    if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $font[2]) { throw "Font hash mismatch: $path" }
    [ordered]@{ File=$font[0]; Url=$url; Sha256=$font[2]; License='SIL Open Font License 1.1'; ProductPayload=$false }
}
Invoke-WebRequest "https://raw.githubusercontent.com/notofonts/noto-fonts/$commit/LICENSE" -OutFile (Join-Path $destination 'OFL.txt')
$manifest | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $destination 'sources.json') -Encoding utf8
Write-Host "Verified test fonts: $destination"
