# node-sharepoint-experiments

Adventures in uploading/listing/downloading documents in Microsoft **SharePoint Online**... using Node.js

## <a name="gettingStarted"></a>Getting Started

```
const Sharepoint = require('@wmfs/sharepoint')
const sp = new Sharepoint('URL HERE')

sp.authenticate()
sp.getWebEndpoint()
sp.getContents(dirPath)
sp.createFolder(options) // options = { dirPath, folderName }
sp.deleteFolder(options) // options = { dirPath, folderName }
sp.createFile(options) // options = { dirPath, fileName, data }
sp.deleteFile(options) // options = { dirPath, fileName }
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
[MIT](https://github.com/wmfs/node-sharepoint-experiments/blob/master/LICENSE)
