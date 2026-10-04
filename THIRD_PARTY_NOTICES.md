# Third-party notices

The application code is licensed under MIT. Third-party components retain their own licenses. Full runtime license texts are also embedded in the standalone HTML.

| Component | Use | License | Source |
| --- | --- | --- | --- |
| SheetJS Community Edition 0.20.3 | Excel/CSV import and XLSX export, vendored standalone script | Apache-2.0 | https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js |
| React / React DOM / Scheduler | User interface | MIT | https://github.com/facebook/react |
| Vite / TypeScript / Vitest / React plugin | Development and build tools | See installed package licenses | Listed in package.json and package-lock.json |

SheetJS was retrieved from the official CDN on 2026-10-03. The unmodified script SHA-256 is `cc015130aa8521e7f088f88898eba949ccdcbfb38df0bd129b44b7273c3a6f41`. Its original notices remain in the script; the full Apache license is in `public/vendor/LICENSE.sheetjs`. React runtime license is in `public/vendor/LICENSE.react`.

## SheetJS required attribution

The following notice is required by the [SheetJS Community Edition license guidance](https://docs.sheetjs.com/docs/miscellany/license/#required-attribution):

```text
SheetJS Community Edition -- https://sheetjs.com/
Copyright (C) 2012-present   SheetJS LLC
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at
      http://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```

Documentation: https://docs.sheetjs.com/docs/getting-started/installation/standalone/ . No runtime CDN request is made by this application. Any vendor upgrade should update its version, license, hash, and import/export regression checks together.
