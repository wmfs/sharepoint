const spauth = require('node-sp-auth')
const axios = require('axios')

class Sharepoint {
  constructor (url) {
    if (!url) {
      throw new Error('You must provide a url.')
    }

    this.url = url
    this.headers = null
    this.site = null
    this.formDigestValue = null
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

  ensureFormDigestValue () {
    if (!this.formDigestValue) {
      this.getFormDigestValue()
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
    this.formDigestValue = data.d.GetContextWebInformation.FormDigestValue
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

  async createFolder (options) {
    this.checkHeaders()
    this.ensureFormDigestValue()

    const { dirPath, folderName } = options

    if (!folderName) {
      throw new Error('You must provide a folder name.')
    }

    await axios({
      method: 'post',
      url: `${this.url}/_api/web/folders`,
      headers: {
        ...this.headers,
        'content-type': 'application/json;odata=verbose',
        'X-RequestDigest': this.formDigestValue
      },
      data: {
        __metadata: { type: 'SP.Folder' },
        ServerRelativeUrl: `${this.site.serverRelativeUrl}${dirPath}/${folderName}`
      },
      responseType: 'json'
    })
  }

  async deleteFolder (options) {
    this.checkHeaders()
    this.ensureFormDigestValue()

    const { dirPath, folderName } = options

    if (!folderName) {
      throw new Error('You must provide a folder name.')
    }

    await axios({
      method: 'post',
      url: `${this.url}/_api/web/GetFolderByServerRelativeUrl('${this.site.serverRelativeUrl}${dirPath}/${folderName}')`,
      headers: {
        ...this.headers,
        'X-RequestDigest': this.formDigestValue,
        'X-HTTP-Method': 'DELETE'
      }
    })
  }

  async createFile (options) {
    this.checkHeaders()
    this.ensureFormDigestValue()

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
        'X-RequestDigest': this.formDigestValue
      }
    })
  }

  async deleteFile (options) {
    this.checkHeaders()
    this.ensureFormDigestValue()

    const { dirPath, fileName } = options

    if (!fileName) {
      throw new Error('You must provide a file name.')
    }

    await axios({
      method: 'post',
      url: `${this.url}/_api/web/GetFileByServerRelativeUrl('${this.site.serverRelativeUrl}${dirPath}/${fileName}')`,
      headers: {
        ...this.headers,
        'X-RequestDigest': this.formDigestValue,
        'X-HTTP-Method': 'DELETE'
      }
    })
  }
}

module.exports = Sharepoint