# sharepoint

[![Tymly Package](https://img.shields.io/badge/tymly-package-blue.svg)](https://tymly.io/)
[![npm (scoped)](https://img.shields.io/npm/v/@wmfs/sharepoint.svg)](https://www.npmjs.com/package/@wmfs/sharepoint)
[![CircleCI](https://circleci.com/gh/wmfs/sharepoint.svg?style=svg)](https://circleci.com/gh/wmfs/sharepoint)
[![CodeFactor](https://www.codefactor.io/repository/github/wmfs/sharepoint/badge)](https://www.codefactor.io/repository/github/wmfs/sharepoint)
[![Dependabot badge](https://img.shields.io/badge/Dependabot-active-brightgreen.svg)](https://dependabot.com/)
[![Commitizen friendly](https://img.shields.io/badge/commitizen-friendly-brightgreen.svg)](http://commitizen.github.io/cz-cli/)
[![JavaScript Style Guide](https://img.shields.io/badge/code_style-standard-brightgreen.svg)](https://standardjs.com)
[![license](https://img.shields.io/github/license/mashape/apistatus.svg)](https://github.com/wmfs/sharepoint/blob/master/README.md)

> A library that allows Node.js applications to interact with a Sharepoint Online site

## General Information
Microsoft supports several authorisation grants and associated token flows.  This library makes use of a 'Client credentials' authentication flow, which permits a 'confidential client' to use its own credentials instead of impersonating a user.

For more information, see https://learn.microsoft.com/en-gb/entra/identity-platform/msal-authentication-flows#client-credentials

Interaction with Sharepoint Online is via the Sharepoint REST API.  For more information, see https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service?tabs=csom and https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-rest-reference/dn450841(v=office.15)

Note that this library makes use of a certificate as a credential, which is required by the Sharepoint REST API.  For future reference, be aware that attempting to use a shared secret instead will result in the Sharepoint REST API refusing access with a HTTP 401 (Unauthorised) error.

## <a name="gettingStarted"></a>Getting Started

### Generating a Certificate/Key/Fingerprint
You will first need to generate a self-signed certificate (and key) using openssl.  Enter the following command into a terminal window to generate a certificate valid for 365 days...

```
openssl req -x509 -newkey rsa:4096 -keyout key.pem -out cert.pem -days 365
```

Note that (if your on Windows), if openssl hangs at any point, press ctrl+c a couple of times to return to the command prompt and re-run the command, but pre-pend '<code>winpty</code>' and a space (so the command starts like <code>winpty openssl...</code>).

After executing the above command, you will be asked to enter the following information...

- A passphrase (twice).  Do make a note of this as you will need it later.
- A 2 character country code ('UK')
- A state/province name ('Greater London')
- A locality name ('London')
- An organisation name ('Home Office')
- Common name ('.')
- Email address ('.')

After entering this information, openssl will generate two files - 'key.pem' (the private key) and 'cert.pem' (the certificate).

We also need the 'fingerprint' of the certificate - you can get this by executing the following command...

```
openssl x509 -in cert.pem -noout -fingerprint
```

This command will output something like this...

> SHA1 Fingerprint=XX:XX:XX:XX:XX:XX:XX:XX:XX:XX:XX:XX:XX:XX:XX:XX:XX:XX:XX:XX

Make a note of the fingerprint, but *get rid of the colons* - in your notes, the fingerprint should be exactly 40 character long.


### Register Application in Azure Portal

1. Sign in to the Microsoft Entra admin centre as at least a 'Cloud Application Administrator'.
2. If you have access to multiple tenants, click Settings in the top menu to switch to the tenant in which you want to register the app from the Directories+subscriptions menu.
3. Browse to Identity > Applications > App registrations and select New registration.
4. In Name, enter 'Test Application'
5. For sign-in audience, select 'Accounts in this organizational directory only'
6. Click the Register button
7. Copy the 'Application (client) ID' and the 'Directory (tenant) ID' and make a note of them somewhere as you'll need them later.
8. On the left, select 'Certificates & secrets'
9. In the 'Certificates' section, select 'Upload certificate'
10. Select the 'cert.pem' file you created earlier and click the Upload button
11. In the description, enter 'Test Application Certificate'
12. Click the Add button.
13. On the left, select 'API Permissions'
14. On the right, under Configured permissions, click the 'Add a permission' button
15. Ensure that the Microsoft APIs tab is selected
16. In the 'Commonly used Microsoft APIs' section, select 'Sharepoint'
17. In the 'Application permissions' section, select the Sites.Selected in the list (use the search box if necessary)
18. Click the 'Add permissions' button at the bottom
19. Click the 'Grant admin consent' button (next to the 'Add a permission' button)
20. In a Web Browser, open Graph Explorer (https://developer.microsoft.com/en-us/graph/graph-explorer) and log in as someone who has the necessary privileges to create/modify sharepoint access permissions.
21. Execute the following GET request, replacing '<site-name>' with the name of your site, and leaving the body empty...

```
https://graph.microsoft.com/v1.0/sites?select=webUrl,Title,Id&$search="<site-name>"
```

You should see something like this - make a note of the value.id property (the 'site id') as you'll need it later...

    {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#sites",
        "value": [
            {
                "id": "XXX.sharepoint.com,XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX,XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX",
                "webUrl": "https://XXX.sharepoint.com/sites/SiteName",
                "displayName": "SiteName"
            }
        ]
    }

22. We now need to make a POST request (though this time we need to put something in the body) to give the above application permissions to read and write to the site, via an URL like this (replacing '<site-id-from-above-GET-request>' with the 'site id' you made a note of above)...

```
https://graph.microsoft.com/v1.0/sites/<site-id-from-above-GET-request>/permissions
```

So for the above site for example, the URL would look like this...

> https://graph.microsoft.com/v1.0/sites/XXX.sharepoint.com,XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX,XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX/permissions

...with this as the body, changing XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX to the 'application (client) id' you made a note of earlier in step 7...

    {
        "roles": [
            "manage"
        ],
        "grantedToIdentities": [
            {
                "application": {
                    "id": "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX",
                    "displayName": "Test Application"
                }
            }
        ]
    }


## <a name="usage"></a>Usage
To use the library, you will need to generate the certificate and key, determine the certificate fingerprint, register your application in Azure Portal and then set up the following environment variables...

| Name                               | Value                                                                                                                                                                                                             |
|------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `SHAREPOINT_AUTH_SCOPE`            | `https://XXX.sharepoint.com/.default`, replacing <code>XXX</code> with your tenant name                                                                                                                           |
| `SHAREPOINT_CLIENT_ID`             | The 'application (client) id' produced by Azure Portal when the application is registered                                                                                                                         |
| `SHAREPOINT_CERT_PRIVATE_KEY_FILE` | The path and filename of the certificate private key file (the 'key.pem' file)                                                                                                                                    |
| `SHAREPOINT_CERT_PASSPHRASE`       | The certificate passphrase                                                                                                                                                                                        |
| `SHAREPOINT_CERT_FINGERPRINT`      | The 40 character (no colons!) hexadecimal fingerprint of the certificate                                                                                                                                          |
| `SHAREPOINT_DEBUG`                 | <code>Y</code> or <code>YES</code> or <code>TRUE</code> if you want information helpful for debugging logged to the console (optional)                                                                            |
| `SHAREPOINT_TENANT_ID`             | The 'directory (tenant) id' produced by Azure Portal when the application is registered                                                                                                                           |

Alternatively, you can edit a `/.env` file if you prefer (as per [dotenv](https://www.npmjs.com/package/dotenv))

Here are the functions the Sharepoint class makes available...

```javascript
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

## <a name="test"></a>Tests
Note that prior to running the tests, you will need to generate the certificate and key, determine the certificate fingerprint, register your application in Azure Portal and then setup the above environment variables.  You will also need to set up the following additional environment variables... 


| Name                               | Value                                                                                                                                                                                                                 |
|------------------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `SHAREPOINT_TESTS_DIR_PATH`        | The location in your site under which the tests will work with when creating/deleting files/folders e.g. `/Shared Documents/General/test`                                                                             |
| `SHAREPOINT_URL`                   | This url of the site that the tests will interact with - something like `https://XXX.sharepoint.com/sites/SiteName`, replacing <code>XXX</code> with your tenant name and <code>SiteName</code> with your site name   |
Again, you can edit a `/.env` file if you prefer (as per [dotenv](https://www.npmjs.com/package/dotenv))

Then, run:
```
npm run test
```

## <a name="license"></a>License
[MIT](https://github.com/wmfs/sharepoint/blob/master/LICENSE)
