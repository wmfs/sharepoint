# sharepoint

[![Tymly Package](https://img.shields.io/badge/tymly-package-blue.svg)](https://tymly.io/)
[![npm (scoped)](https://img.shields.io/npm/v/@wmfs/sharepoint.svg)](https://www.npmjs.com/package/@wmfs/sharepoint)
[![CircleCI](https://circleci.com/gh/wmfs/sharepoint.svg?style=svg)](https://circleci.com/gh/wmfs/sharepoint)
[![CodeFactor](https://www.codefactor.io/repository/github/wmfs/sharepoint/badge)](https://www.codefactor.io/repository/github/wmfs/sharepoint)
[![Dependabot badge](https://img.shields.io/badge/Dependabot-active-brightgreen.svg)](https://dependabot.com/)
[![Commitizen friendly](https://img.shields.io/badge/commitizen-friendly-brightgreen.svg)](http://commitizen.github.io/cz-cli/)
[![JavaScript Style Guide](https://img.shields.io/badge/code_style-standard-brightgreen.svg)](https://standardjs.com)
[![license](https://img.shields.io/github/license/mashape/apistatus.svg)](https://github.com/wmfs/sharepoint/blob/master/README.md)

> Adventures in uploading/listing/downloading documents in Microsoft **SharePoint Online**... using Node.js

## <a name="gettingStarted"></a>Getting Started

```
const Sharepoint = require('@wmfs/sharepoint')
const sp = new Sharepoint('URL HERE')

sp.authenticate()
sp.getWebEndpoint()
sp.getContents(path)
sp.createFolder(path)
sp.deleteFolder(path)
sp.createFile(options) // options = { path, fileName, data }
sp.deleteFile(options) // options = { path, fileName }
sp.createFileChunked(options) // options = { path, fileName, stream, fileSize, chunkSize }
```

## <a name="test"></a>Test
First, set these to match your SharePoint environment:

| Env Variable | Value |
| ------------ | ----- |
| `SHAREPOINT_URL` | This is the site we're aiming for, so something like `https://example.sharepoint.com/sites/YourSite/` |
| `SHAREPOINT_USERNAME` | The username you want to connect to SharePoint with. Note this is the full username with an `@`, so something like `some.username@example.com` |
| `SHAREPOINT_PASSWORD` | And yup, the password to accompany `SHAREPOINT_USERNAME`. |
| `SHAREPOINT_DIR_PATH` | Path to where the files are. e.g. `/Shared Documents/General ` |

* Alternatively, you can edit a `/.env` file if you prefer (as per [dotenv](https://www.npmjs.com/package/dotenv))

Then, run:
```
npm run test
```

## <a name="license"></a>License
[MIT](https://github.com/wmfs/sharepoint/blob/master/LICENSE)


