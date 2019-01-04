const spauth = require('node-sp-auth')
const axios = require('axios')

const getGUID = () => {
  let d = Date.now()
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = (d + Math.random() * 16) % 16 | 0
    d = Math.floor(d / 16)
    return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16)
  })
}

class Sharepoint {
  constructor (url) {
    if (!url) {
      throw new Error('You must provide a url.')
    }

    this.url = url
    this.headers = null
    this.site = null
  }

  async authenticate (username, password) {
    if (!username && !password) {
      throw new Error('You must provide a username and password.')
    }

    const { headers } = await spauth.getAuth(this.url, { username, password })
    this.headers = {
      ...headers,
      Accept: 'application/json;odata=verbose'
    }
  }

  checkHeaders () {
    if (!this.headers) {
      throw new Error('No headers, you must authenticate.')
    }
  }

  async getWebEndpoint () {
    this.checkHeaders()

    const { url, headers } = this

    const { data } = await axios.get(
      `${url}/_api/web`,
      { headers, responseType: 'json' }
    )

    const site = data.d

    this.site = {
      id: site.Id,
      title: site.Title,
      description: site.Description,
      created: site.Created,
      serverRelativeUrl: site.ServerRelativeUrl,
      lastModified: site.LastItemUserModifiedDate
    }
  }

  async getFormDigestValue () {
    this.checkHeaders()

    const { data } = await axios({
      method: 'post',
      url: `${this.url}/_api/contextinfo`,
      headers: {
        ...this.headers,
        'content-type': 'application/json;odata=verbose'
      },
      responseType: 'json'
    })

    return data.d.GetContextWebInformation.FormDigestValue
  }

  async getContents (dirPath) {
    this.checkHeaders()

    const { url, headers, site } = this

    const get = type => {
      return axios.get(
        `${url}/_api/web/GetFolderByServerRelativeUrl('${site.serverRelativeUrl}${dirPath}')/${type}`,
        { headers, responseType: 'json' }
      )
    }

    const folders = await get('Folders')
    const files = await get('Files')

    return [...folders.data.d.results, ...files.data.d.results]
  }

  async createFolder (path) {
    this.checkHeaders()
    const formDigestValue = await this.getFormDigestValue()

    if (!path) {
      throw new Error('You must provide a path.')
    }

    await axios({
      method: 'post',
      url: `${this.url}/_api/web/folders`,
      headers: {
        ...this.headers,
        'content-type': 'application/json;odata=verbose',
        'X-RequestDigest': formDigestValue
      },
      data: {
        __metadata: { type: 'SP.Folder' },
        ServerRelativeUrl: `${this.site.serverRelativeUrl}${path}`
      },
      responseType: 'json'
    })
  }

  async deleteFolder (path) {
    this.checkHeaders()
    const formDigestValue = await this.getFormDigestValue()

    if (!path) {
      throw new Error('You must provide a path.')
    }

    await axios({
      method: 'post',
      url: `${this.url}/_api/web/GetFolderByServerRelativeUrl('${this.site.serverRelativeUrl}${path}')`,
      headers: {
        ...this.headers,
        'X-RequestDigest': formDigestValue,
        'X-HTTP-Method': 'DELETE'
      }
    })
  }

  async createFile (options) {
    this.checkHeaders()
    const formDigestValue = await this.getFormDigestValue()

    const { dirPath, fileName, data } = options

    if (!fileName) {
      throw new Error('You must provide a file name.')
    }

    if (!data) {
      throw new Error('You must provide data.')
    }

    await axios({
      method: 'post',
      url: `${this.url}/_api/web/GetFolderByServerRelativeUrl('${this.site.serverRelativeUrl}${dirPath}')/Files/add(url='${fileName}', overwrite=true)`,
      data,
      headers: {
        ...this.headers,
        'X-RequestDigest': formDigestValue
      }
    })
  }

  async createFileChunked (options) {
    const {
      dirPath, fileName, stream, fileSize
    } = options

    const chunkSize = options.chunkSize || 65536
    this.checkHeaders()
    const formDigestValue = await this.getFormDigestValue()

    await this.createFile({
      dirPath,
      fileName,
      data: ' '
    })

    const uploadId = getGUID()

    let firstChunk = true
    let sent = 0
    const self = this

    const upload = new Promise(function (resolve, reject) {
      stream.on('data', async (data) => {
        try {
          stream.pause()
          if (firstChunk) {
            firstChunk = false
            const response = await axios({
              method: 'post',
              url: `${self.url}/_api/web/getfilebyserverrelativeurl('${self.site.serverRelativeUrl}${dirPath}/${fileName}')/startupload(uploadId=guid'${uploadId}')`,
              data,
              headers: {
                ...self.headers,
                'X-RequestDigest': formDigestValue
              }
            })
            sent = Number(response.data.d.StartUpload)
          } else if (sent + chunkSize >= fileSize) {
            await axios({
              method: 'post',
              url: `${self.url}/_api/web/getfilebyserverrelativeurl('${self.site.serverRelativeUrl}${dirPath}/${fileName}')/finishupload(uploadId=guid'${uploadId}',fileoffset=${sent})`,
              data,
              headers: {
                ...self.headers,
                'X-RequestDigest': formDigestValue
              }
            })
            resolve()
          } else {
            const response = await axios({
              method: 'post',
              url: `${self.url}/_api/web/getfilebyserverrelativeurl('${self.site.serverRelativeUrl}${dirPath}/${fileName}')/continueupload(uploadId=guid'${uploadId}',fileoffset=${sent})`,
              data,
              headers: {
                ...self.headers,
                'X-RequestDigest': formDigestValue
              }
            })
            sent = Number(response.data.d.ContinueUpload)
          }
          stream.resume()
        } catch (e) {
          stream.destroy()
          await axios({
            method: 'post',
            url: `${self.url}/_api/web/getfilebyserverrelativeurl('${self.site.serverRelativeUrl}${dirPath}/${fileName}')/cancelupload(uploadId=guid'${uploadId}')`,
            headers: {
              ...self.headers,
              'X-RequestDigest': formDigestValue
            }
          })
          reject(e)
        }
      })

      stream.on('error', async err => {
        await axios({
          method: 'post',
          url: `${self.url}/_api/web/getfilebyserverrelativeurl('${self.site.serverRelativeUrl}${dirPath}/${fileName}')/cancelupload(uploadId=guid'${uploadId}')`,
          headers: {
            ...self.headers,
            'X-RequestDigest': formDigestValue
          }
        })
        reject(err)
      })
    })

    await upload
  }

  async deleteFile (options) {
    this.checkHeaders()
    const formDigestValue = await this.getFormDigestValue()

    const { dirPath, fileName } = options

    if (!fileName) {
      throw new Error('You must provide a file name.')
    }

    await axios({
      method: 'post',
      url: `${this.url}/_api/web/GetFileByServerRelativeUrl('${this.site.serverRelativeUrl}${dirPath}/${fileName}')`,
      headers: {
        ...this.headers,
        'X-RequestDigest': formDigestValue,
        'X-HTTP-Method': 'DELETE'
      }
    })
  }
}

module.exports = Sharepoint
