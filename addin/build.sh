#!/bin/sh
# Build the add-in DLL and the TestHost with the C# compiler (mono csc on macOS, or csc.exe on Windows).
# Targets .NET Framework 4.8 reference assemblies, AnyCPU (loads in 64-bit and 32-bit Outlook).
set -eu
cd "$(dirname "$0")"
REF="${NETFX_REF:-/opt/homebrew/Cellar/mono/6.14.1/lib/mono/4.8-api}"
OUT="${1:-../build/addin}"
mkdir -p "$OUT"
REFS="-nostdlib -r:$REF/mscorlib.dll -r:$REF/System.dll -r:$REF/System.Core.dll -r:$REF/System.Drawing.dll \
 -r:$REF/System.Windows.Forms.dll -r:$REF/System.Web.Extensions.dll -r:$REF/Microsoft.CSharp.dll -r:$REF/System.IO.Compression.dll"
csc -nologo -deterministic -langversion:7.3 -target:library -platform:anycpu -optimize+ -warnaserror- -nowarn:1591 \
    $REFS -out:"$OUT/HatirMail.dll" HatirMail.cs SearchForm.cs
csc -nologo -deterministic -langversion:7.3 -target:winexe -platform:anycpu -optimize+ \
    $REFS -r:$REF/System.Xml.dll -r:$REF/System.Xml.Linq.dll -r:$REF/Accessibility.dll -r:"$OUT/HatirMail.dll" -out:"$OUT/HatirMail.TestHost.exe" TestHost.cs
echo "built: $OUT/HatirMail.dll $OUT/HatirMail.TestHost.exe"
csc -nologo -deterministic -langversion:7.3 -target:exe -platform:anycpu -optimize+ \
    $REFS -out:"$OUT/HatirMail.OlbCheck.exe" OlbCheck.cs
