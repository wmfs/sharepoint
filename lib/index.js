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
   * @constructor
   * @param siteUrl a tenant based url to the site whose file system you wish to interact with.  For example https://example.sharepoint.com/sites/SiteName (replacing 'example' with your tenant name and 'SiteName' with your site name).
   */
  constructor (siteUrl) {
    if (!siteUrl) {
      throw new Error('siteUrl has not been specified')
    }

    const authScope = process.env.SHAREPOINT_AUTH_SCOPE
    if (!authScope) {
      throw new Error('SHAREPOINT_AUTH_SCOPE environment variable has not been set')
    }

    if (!(authScope.toLowerCase().startsWith('https://') && authScope.toLowerCase().endsWith('.sharepoint.com/.default'))) {
      throw new Error('SHAREPOINT_AUTH_SCOPE environment variable value is not valid - it must begin with "https://" and end with ".sharepoint.com/.default"')
    }

    const clientId = process.env.SHAREPOINT_CLIENT_ID
    if (!clientId) {
      throw new Error('SHAREPOINT_CLIENT_ID environment variable has not been set')
    }

    const tenantId = process.env.SHAREPOINT_TENANT_ID
    if (!tenantId) {
      throw new Error('SHAREPOINT_TENANT_ID environment variable has not been set')
    }

    const certPassphrase = process.env.SHAREPOINT_CERT_PASSPHRASE
    if (!certPassphrase) {
      throw new Error('SHAREPOINT_CERT_PASSPHRASE environment variable has not been set')
    }

    const certFingerprint = process.env.SHAREPOINT_CERT_FINGERPRINT
    if (!certFingerprint) {
      throw new Error('SHAREPOINT_CERT_FINGERPRINT environment variable has not been set')
    }

    if (certFingerprint.length !== 40) {
      throw new Error('SHAREPOINT_CERT_FINGERPRINT environment variable value is not valid - it must be exactly 40 characters in length')
    }

    const certPrivateKeyFileFile = process.env.SHAREPOINT_CERT_PRIVATE_KEY_FILE
    if (!certPrivateKeyFileFile) {
      throw new Error('SHAREPOINT_CERT_PRIVATE_KEY_FILE environment variable has not been set')
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

  /**
   * Determines the base path of the site and populates the baseUrl and encodedBaseUrl properties.
   * So for example, if your site url is 'https://example.sharepoint.com/sites/TestSite', then your base url would be '/sites/TestSite'.
   * This is used to construct paths when interacting with the sites file system.
   * @returns {Promise<void>}
   */
  async getWebEndpoint () {
    checkHeaders(this.accessToken)

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
      logAxiosError(this.debug, err, 'Unable to get web endpoint')
    }

    this.baseUrl = response.data.d.ServerRelativeUrl
    this.encodedBaseUrl = encodeURIComponent(response.data.d.ServerRelativeUrl)
  }

  /**
   * Returns an array of objects, each describing a file or folder in the specified folder.
   * Note that folders will appear in the array first, and both files and folders will be
   * sorted by name.
   * @param path The path representing a folder relative to the site folder.
   * @returns {Promise<*[]>} An array of objects, each describing a file or folder
   */
  async getContents (path) {
    checkHeaders(this.accessToken)

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
        logAxiosError(this.debug, err, `Failed to get folder contents (type: ${type})`)
      }
    }

    const folders = await get('Folders')
    const files = await get('Files')

    // natural sort of files/folders
    const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' })
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

    checkHeaders(this.accessToken)

    const formDigestValue = await getFormDigestValue(this.siteUrl, this.accessToken)
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
      logAxiosError(this.debug, err, 'Failed to create specified folder')
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

    checkHeaders(this.accessToken)

    const formDigestValue = await getFormDigestValue(this.siteUrl, this.accessToken)
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
      logAxiosError(this.debug, err, 'Unable to delete folder')
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

    checkHeaders(this.accessToken)

    const { data } = options
    const path = encodeURIComponent(options.path)
    const fileName = encodeURIComponent(options.fileName)

    const formDigestValue = await getFormDigestValue(this.siteUrl, this.accessToken)
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
      logAxiosError(this.debug, err, 'Unable to create file')
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

    checkHeaders(this.accessToken)

    const formDigestValue = await getFormDigestValue(this.siteUrl, this.accessToken)
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

    checkHeaders(this.accessToken)

    const path = encodeURIComponent(options.path)
    const fileName = encodeURIComponent(options.fileName)

    const formDigestValue = await getFormDigestValue(this.siteUrl, this.accessToken)
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
      logAxiosError(this.debug, err, 'Unable to delete file')
    }
  }

  /**
   * Moves the specified file
   * @param options An object that must contain a 'fileName' (the name of the file), 'sourcePath' (the path to a folder in
   * which the file is to be moved from) and 'targetPath' (the path to a folder which the file is to be moved to) properties.
   * @returns {Promise<void>}
   */
  async moveFile (options) {
    if (!options.fileName) {
      throw new Error('You must provide a file name.')
    }

    checkHeaders(this.accessToken)

    const sourcePath = encodeURIComponent(options.sourcePath)
    const targetPath = encodeURIComponent(options.targetPath)
    const fileName = encodeURIComponent(options.fileName)

    const formDigestValue = await getFormDigestValue(this.siteUrl, this.accessToken)

    const url = `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${this.encodedBaseUrl}${sourcePath}/${fileName}')/moveto(newurl='${this.encodedBaseUrl}${targetPath}/${fileName}',flags=1)`

    try {
      await axios({
        method: 'post',
        url,
        headers: {
          Authorization: `Bearer ${this.accessToken}`,
          'X-RequestDigest': formDigestValue
        }
      })
    } catch (err) {
      console.log(err)
      logAxiosError(this.debug, err, 'Unable to move file')
    }
  }

  /**
   * Moves the specified file
   * @param options An object that must contain a 'sourcePath' (the path to a folder in
   * which the file is to be moved from) and 'targetPath' (the path to a folder which the file is to be moved to) properties.
   * @returns {Promise<void>}
   */
  async moveFolder (options) {
    checkHeaders(this.accessToken)

    const filesToMove = []

    async function collectFilesToMove (self, sourcePath, targetPath) {
      const contents = await self.getContents(sourcePath)
      for (const item of contents) {
        if (item.__metadata.type === 'SP.File') {
          filesToMove.push({ sourcePath, targetPath, fileName: item.Name })
        } else if (item.__metadata.type === 'SP.Folder') {
          await collectFilesToMove(self, `${sourcePath}/${item.Name}`, `${targetPath}/${item.Name}`)
        }
      }
    }

    await collectFilesToMove(this, options.sourcePath, options.targetPath)

    for (const { sourcePath, targetPath, fileName } of filesToMove) {
      await this.createFolder(targetPath)
      await this.moveFile({ sourcePath, targetPath, fileName })
    }

    await this.deleteFolder(options.sourcePath)
  }
}

// based on https://axios-http.com/docs/handling_errors
function logAxiosError (debug, err, msg) {
  if (debug) {
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

function checkHeaders (accessToken) {
  if (!accessToken) {
    throw new Error('Access token not available - please authenticate() prior to calling this function')
  }
}

async function getFormDigestValue (siteUrl, accessToken) {
  checkHeaders(accessToken)

  let response
  try {
    response = await axios({
      method: 'post',
      url: `${siteUrl}/_api/contextinfo`,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: 'application/json;odata=verbose'
      },
      responseType: 'json',
      data: {}
    })
  } catch (err) {
    logAxiosError(err, 'Unable to get form digest value')
  }

  return response.data.d.GetContextWebInformation.FormDigestValue
}

module.exports = Sharepoint
