# node-sharepoint-experiments

Adventures in uploading/listing/downloading documents in Microsoft **SharePoint Online**... using Node.js

## Environment Variables

First, set these to match your SharePoint environment:

| Env Variable | Value |
| ------------ | ----- |
| `SHAREPOINT_URL` | This is the site we're aiming for, so something like `https://example.sharepoint.com/sites/YourSite/` |
| `SHAREPOINT_USERNAME` | The username you want to connect to SharePoint with. Note this is the full username with an `@`, so something like `some.username@example.com` |
| `SHAREPOINT_PASSWORD` | And yup, the password to accompany `SHAREPOINT_USERNAME` so `Shhh!` or whatever. |

* Alternatively, you can edit a `/.env` file if you prefer (as per [dotenv](https://www.npmjs.com/package/dotenv))

## <a name="running"></a>Running

```
npm install
npm start
```

## <a name="license"></a>License
[MIT](https://github.com/wmfs/node-sharepoint-experiments/blob/master/LICENSE)
