const process = require('node:process')
const crypto = require('node:crypto')
const fs = require('node:fs')
const msal = require('@azure/msal-node')
const axios = require('axios')
const { v4: uuid } = require('uuid')

// see https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service?tabs=csom
class Sharepoint {
  /**
   * Sets up an instance of the Sharepoint class to interact with the filesystem under the specified site url.
   * Note that you must have set up the following environment variables in your run environment, or it will throw an Error.
   * | name                             | description                                                                                     |
   * | -------------------------------- | ----------------------------------------------------------------------------------------------- |
   * | SHAREPOINT_AUTH_SCOPE            | 'https://example.sharepoint.com/.default', replacing <code>example</code> with your tenant name |
   * | SHAREPOINT_CERT_FINGERPRINT      | The 40 character (no colons!) hexadecimal fingerprint of the certificate                        |
   * | SHAREPOINT_CERT_PASSPHRASE       | The certificate passphrase                                                                      |
   * | SHAREPOINT_CERT_PRIVATE_KEY_FILE | The path and filename of the certificate private key file                                       |
   * | SHAREPOINT_CLIENT_ID             | The 'application (client) id' produced by Azure Portal when the application is registered       |
   * | SHAREPOINT_TENANT_ID             | The 'directory (tenant) id' produced by Azure Portal when the application is registered         |
   * @constructor
   * @param siteUrl a tenant based url to the site whose file system you wish to interact with.  For example https://example.sharepoint.com/sites/SiteName (replacing example with your tenant name and SiteName with your site name).
   */
  constructor (siteUrl) {
    if (!siteUrl) {
      throw new Error('sharepoint site url has not been specified')
    }

    const authScope = process.env.SHAREPOINT_AUTH_SCOPE
    if (!authScope) {
      throw new Error('sharepoint authentication scope has not been specified')
    }

    if (!(authScope.toLowerCase().startsWith('https://') && authScope.toLowerCase().endsWith('.sharepoint.com/.default'))) {
      throw new Error('Specified authentication scope is not valid - it must begin with "https://" and end with ".sharepoint.com/.default"')
    }

    const clientId = process.env.SHAREPOINT_CLIENT_ID
    if (!clientId) {
      throw new Error('sharepoint client id has not been specified')
    }

    const tenantId = process.env.SHAREPOINT_TENANT_ID
    if (!tenantId) {
      throw new Error('sharepoint tenant id has not been specified')
    }

    const certPassphrase = process.env.SHAREPOINT_CERT_PASSPHRASE
    if (!certPassphrase) {
      throw new Error('sharepoint certificate passphrase has not been specified')
    }

    const certFingerprint = process.env.SHAREPOINT_CERT_FINGERPRINT
    if (!certFingerprint) {
      throw new Error('sharepoint certificate fingerprint has not been specified')
    }

    if (certFingerprint.length !== 40) {
      throw new Error('sharepoint certificate fingerprint is not 40 characters in length')
    }

    const certPrivateKeyFileFile = process.env.SHAREPOINT_CERT_PRIVATE_KEY_FILE
    if (!certPrivateKeyFileFile) {
      throw new Error('sharepoint certificate private key file path/filename has not been specified')
    }

    if (!(fs.existsSync(certPrivateKeyFileFile) && fs.lstatSync(certPrivateKeyFileFile).isFile())) {
      throw new Error(`specified sharepoint certificate private key file ('${certPrivateKeyFileFile}') does not exist`)
    }

    this.siteUrl = siteUrl
    this.authScope = authScope
    this.accessToken = null
    this.baseUrl = null
    this.encodedBaseUrl = null

    const certPrivateKeyObject = crypto.createPrivateKey({
      key: fs.readFileSync(certPrivateKeyFileFile, 'utf8'),
      passphrase: certPassphrase,
      format: 'pem'
    })
    const certPrivateKey = certPrivateKeyObject.export({
      format: 'pem',
      type: 'pkcs8'
    })

    const config = {
      auth: {
        clientCertificate: {
          thumbprint: certFingerprint,
          privateKey: certPrivateKey
        },
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`
      }
    }

    this.debug = process.env.SHAREPOINT_DEBUG
    this.debug = this.debug && (this.debug.toUpperCase() === 'Y' || this.debug.toUpperCase() === 'YES' || this.debug.toUpperCase() === 'TRUE')
    if (this.debug) {
      // configure msal to log debugging information to the console
      config.system = {
        loggerOptions: {
          loggerCallback (loglevel, message, containsPii) {
            console.log(message)
          },
          piiLoggingEnabled: false,
          logLevel: msal.LogLevel.Verbose
        }
      }
    }

    this.cca = new msal.ConfidentialClientApplication(config)
  }

  /**
   * Carries out the login process and then internally stores the access token, which is used when interacting with the Sharepoint REST API.
   * @returns {Promise<void>}
   */
  async authenticate () {
    const { accessToken } = await this.cca.acquireTokenByClientCredential({
      scopes: [this.authScope]
    })
    this.accessToken = accessToken
  }

  checkHeaders () {
    if (!this.accessToken) {
      throw new Error('Access token not available - please authenticate() prior to calling this function')
    }
  }

  /**
   * Determines the base path of the site and populates the baseUrl and encodedBaseUrl properties.
   * So for example, if your site url is 'https://example.sharepoint.com/sites/TestSite', then your base url would be '/sites/TestSite'.
   * This is used to construct paths when interacting with the sites file system.
   * @returns {Promise<void>}
   */
  async getWebEndpoint () {
    this.checkHeaders()

    let response
    try {
      response = await axios.get(
        `${this.siteUrl}/_api/web`,
        {
          headers: {
            Authorization: `Bearer ${this.accessToken}`,
            Accept: 'application/json;odata=verbose'
          }
        }
      )
    } catch (err) {
      this.logAxiosError(err, 'Unable to get web endpoint')
    }

    this.baseUrl = response.data.d.ServerRelativeUrl
    this.encodedBaseUrl = encodeURIComponent(response.data.d.ServerRelativeUrl)
  }

  async getFormDigestValue () {
    this.checkHeaders()

    let response
    try {
      response = await axios({
        method: 'post',
        url: `${this.siteUrl}/_api/contextinfo`,
        headers: {
          Authorization: `Bearer ${this.accessToken}`,
          Accept: 'application/json;odata=verbose'
        },
        responseType: 'json',
        data: {}
      })
    } catch (err) {
      this.logAxiosError(err, 'Unable to get form digest value')
    }

    return response.data.d.GetContextWebInformation.FormDigestValue
  }

  /**
   * Returns an array of folders, followed by files in the specified folder.
   * @param path The path representing a folder relative to the site folder.
   * @returns {Promise<*[]>} An array of objects, each describing a file or folder
   */
  async getContents (path) {
    this.checkHeaders()

    const get = type => {
      try {
        return axios.get(
          `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${this.encodedBaseUrl}${encodeURIComponent(path)}')/${type}`,
          {
            headers: {
              Authorization: `Bearer ${this.accessToken}`,
              Accept: 'application/json;odata=verbose'
            },
            responseType: 'json'
          }
        )
      } catch (err) {
        this.logAxiosError(err, `Failed to get folder contents (type: ${type})`)
      }
    }

    const folders = await get('Folders')
    const files = await get('Files')

    // natural sort of files/folders
    const collator = new Intl.Collator(undefined, {numeric: true, sensitivity: 'base'})
    const nameCompare = (a, b) => {
      return collator.compare(a.Name, b.Name)
    }

    folders.data.d.results.sort(nameCompare)
    files.data.d.results.sort(nameCompare)

    return [...folders.data.d.results, ...files.data.d.results]
  }

  /**
   * Create the specified folder.
   * @param path The path of the folder you want to create relative to the site folder
   * @returns {Promise<void>}
   */
  async createFolder (path) {
    if (!path) {
      throw new Error('You must provide a path.')
    }

    this.checkHeaders()

    const formDigestValue = await this.getFormDigestValue()
    try {
      await axios({
        method: 'post',
        url: `${this.siteUrl}/_api/web/folders`,
        headers: {
          Authorization: `Bearer ${this.accessToken}`,
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose', // docs say use 'application/json', but call fails if we do
          'X-RequestDigest': formDigestValue
        },
        data: {
          __metadata: { type: 'SP.Folder' },
          ServerRelativeUrl: `${this.encodedBaseUrl}${encodeURIComponent(path)}`
        },
        responseType: 'json'
      })
    } catch (err) {
      this.logAxiosError(err, 'Failed to create specified folder')
    }
  }

  /**
   * Deletes the specified folder.
   * @param path The path of the folder you want to delete relative to the site folder
   * @returns {Promise<void>}
   */
  async deleteFolder (path) {
    if (!path) {
      throw new Error('You must provide a path.')
    }

    this.checkHeaders()

    const formDigestValue = await this.getFormDigestValue()
    try {
      await axios({
        method: 'post',
        url: `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${this.encodedBaseUrl}${encodeURIComponent(path)}')`,
        headers: {
          Authorization: `Bearer ${this.accessToken}`,
          'X-RequestDigest': formDigestValue,
          'X-HTTP-Method': 'DELETE'
        }
      })
    } catch (err) {
      this.logAxiosError(err, 'Unable to delete folder')
    }
  }

  /**
   * Creates and populates a (non-binary) file.  Note that if the specified file already exists, it will be overwritten.
   * @param options An object that must contain a 'fileName' (the name of the file), 'path' (the path to a folder in
   * which the file will be created) and 'data' (the contents of the file) properties.
   * @returns {Promise<void>}
   */
  async createFile (options) {
    if (!options.fileName) {
      throw new Error('You must provide a file name.')
    }

    if (!options.data) {
      throw new Error('You must provide data.')
    }

    this.checkHeaders()

    const { data } = options
    const path = encodeURIComponent(options.path)
    const fileName = encodeURIComponent(options.fileName)

    const formDigestValue = await this.getFormDigestValue()
    try {
      await axios({
        method: 'post',
        url: `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${this.encodedBaseUrl}${path}')/Files/add(url='${fileName}', overwrite=true)`,
        data,
        headers: {
          Accept: 'application/json;odata=verbose',
          Authorization: `Bearer ${this.accessToken}`,
          'X-RequestDigest': formDigestValue
        }
      })
    } catch (err) {
      this.logAxiosError(err, 'Unable to create file')
    }
  }

  // see https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-rest-reference/dn450841(v=office.15)
  /**
   * Creates a file and uploads its contents in chunks.
   * @param options An object that must contain a 'fileName' (the name of the file), 'path' (the path to a folder in
   * which the file will be created), 'stream' (a file data stream) and 'fileSize' (the size of the file in bytes)
   * properties.  It can also optionally specify a 'chunkSize' property to specify the size (again in bytes) of each
   * chunk
   * @returns {Promise<{filePath: string, url: string, Name: string}>}
   */
  async createFileChunked (options) {
    const { stream, fileSize } = options

    const path = encodeURIComponent(options.path)
    const fileName = encodeURIComponent(options.fileName)

    const chunkSize = options.chunkSize || 65536
    this.checkHeaders()
    const formDigestValue = await this.getFormDigestValue()

    await this.createFile({
      path,
      fileName,
      data: ' '
    })

    const uploadId = uuid()

    let firstChunk = true
    let sent = 0
    const self = this
    const baseUploadUrl = `${self.siteUrl}/_api/web/GetFileByServerRelativeUrl('${self.encodedBaseUrl}${path}${encodeURIComponent('/')}${fileName}')`

    await new Promise(function (resolve, reject) {
      stream.on('data', async (data) => {
        try {
          stream.pause()
          if (firstChunk) {
            firstChunk = false
            const response = await axios({
              method: 'post',
              url: `${baseUploadUrl}/startupload(uploadId=guid'${uploadId}')`,
              data,
              headers: {
                Authorization: `Bearer ${self.accessToken}`,
                'X-RequestDigest': formDigestValue
              }
            })
            sent = Number(response.data.value)

            if (sent >= fileSize) {
              await axios({
                method: 'post',
                url: `${baseUploadUrl}/finishupload(uploadId=guid'${uploadId}',fileoffset=${sent})`,
                headers: {
                  Authorization: `Bearer ${self.accessToken}`,
                  'X-RequestDigest': formDigestValue
                }
              })
              resolve()
            }
          } else if (sent + chunkSize >= fileSize) {
            await axios({
              method: 'post',
              url: `${baseUploadUrl}/finishupload(uploadId=guid'${uploadId}',fileoffset=${sent})`,
              data,
              headers: {
                Authorization: `Bearer ${self.accessToken}`,
                'X-RequestDigest': formDigestValue
              }
            })
            resolve()
          } else {
            const response = await axios({
              method: 'post',
              url: `${baseUploadUrl}/continueupload(uploadId=guid'${uploadId}',fileoffset=${sent})`,
              data,
              headers: {
                Authorization: `Bearer ${self.accessToken}`,
                'X-RequestDigest': formDigestValue
              }
            })
            sent = Number(response.data.value)
          }

          stream.resume()
        } catch (e) {
          stream.destroy()
          await axios({
            method: 'post',
            url: `${baseUploadUrl}/cancelupload(uploadId=guid'${uploadId}')`,
            headers: {
              Authorization: `Bearer ${self.accessToken}`,
              'X-RequestDigest': formDigestValue
            }
          })
          reject(e)
        }
      })

      stream.on('error', async err => {
        await axios({
          method: 'post',
          url: `${baseUploadUrl}/cancelupload(uploadId=guid'${uploadId}')`,
          headers: {
            Authorization: `Bearer ${self.accessToken}`,
            'X-RequestDigest': formDigestValue
          }
        })
        reject(err)
      })
    })

    return {
      Name: fileName,
      filePath: `${path}/${fileName}`,
      url: `${this.baseUrl}/${path}/${fileName}`
    }
  }

  /**
   * Deletes the specified file
   * @param options An object that must contain a 'fileName' (the name of the file) and 'path' (the path to a folder in
   * which the file is to be deleted) properties.
   * @returns {Promise<void>}
   */
  async deleteFile (options) {
    if (!options.fileName) {
      throw new Error('You must provide a file name.')
    }

    this.checkHeaders()

    const path = encodeURIComponent(options.path)
    const fileName = encodeURIComponent(options.fileName)

    const formDigestValue = await this.getFormDigestValue()
    try {
      await axios({
        method: 'post',
        url: `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${this.encodedBaseUrl}${path}/${fileName}')`,
        headers: {
          Authorization: `Bearer ${this.accessToken}`,
          'X-RequestDigest': formDigestValue,
          'X-HTTP-Method': 'DELETE'
        }
      })
    } catch (err) {
      this.logAxiosError(err, 'Unable to delete file')
    }
  }

  // based on https://axios-http.com/docs/handling_errors
  logAxiosError (err, msg) {
    if (this.debug) {
      if (err.response) {
        // request was made but server responded with a non-2xx status code
        console.log(`server responded with status code ${err.response.status}`)
        console.log(`data: ${err.response.data}`)
      } else if (err.request) {
        // request was made but no response was received
        console.log(err.request)
      } else {
        // something happened in setting up the request that triggered an error
        console.log('Error', err.message)
      }
      console.log(err.config)
    }
    throw new Error(msg)
  }
}

module.exports = Sharepoint
