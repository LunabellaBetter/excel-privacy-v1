# Third-party notices for Excel Privacy V1.0

The project MIT license covers only the independently maintained V1 source
snapshot listed by release/SOURCE_MANIFEST.json. It does not license the
historical kwen1510 source, legacy CLI, Skill template, or old examples excluded
from this snapshot. No dean9703111 code is included in V1.

Runtime components:

| Component | Version | License |
| --- | --- | --- |
| Python | 3.12.x | PSF and included historical notices |
| openpyxl | 3.1.5 | MIT |
| et-xmlfile | 2.0.0 | MIT |
| cryptography | 50.0.2 | Apache-2.0 OR BSD-3-Clause |
| cffi | 2.1.1 | MIT-0 |
| pycparser | 3.0 | BSD-3-Clause |
| Tcl/Tk | platform build version | Tcl/Tk license terms |
| OpenSSL and Rust dependencies | see bundled cryptography SBOMs | per-component terms |

Build components include PyInstaller 6.22.3 (GPL with application-distribution
exception), community hooks 2026.8 (GPL standard hooks; Apache-2.0 runtime hooks),
altgraph 0.17.5 (MIT), packaging 26.3 (Apache-2.0 OR BSD-2-Clause), macholib
1.16.4 on Mac (MIT), and pefile/pywin32-ctypes on Windows (see native inventory).

Full license and copyright texts are in release/licenses in source packages
and licenses in frozen applications. Platform-specific collected inventories
and SBOMs are authoritative for the corresponding built asset. Development
tools are not automatically application dependencies.

This is an attribution inventory, not a security certification. Public release
also requires review of native platform runtime additions and final artifacts.

The macOS native library audit additionally found OpenSSL 3.6.5 (Python's
stdlib SSL, distinct from cryptography's static OpenSSL 4.0.3), liblzma from
XZ 5.8.4, mpdecimal 4.0.1 and LibTomMath 1.3.0. Their full native license texts
are included under mac-* notice directories; mac-shared-libraries.json records
that inventory. XZ's COPYING distinguishes the library's terms from tools
not distributed here. These version labels describe the Mac build only.

Windows Python runtime binaries additionally include OpenSSL 3.0.16 and zlib
1.3.1; their exact upstream license texts are in windows-openssl-3.0.16 and
windows-zlib-1.3.1. This is distinct from cryptography's own OpenSSL SBOM.
The final Windows ZIP excludes VCRUNTIME140*.dll, ucrtbase.dll and api-ms-win-*
Microsoft runtime binaries. They are provided by Windows 10/11 and the user's
official Microsoft VC++ x64 runtime installation. The package does not grant
or assume permission to redistribute Microsoft runtime binaries.

The delivery-level licenses-supplement directory additionally carries the
CPython 3.12 incorporated-software acknowledgements and libffi notice. These
supplement the application-internal notices without altering signed binaries.
