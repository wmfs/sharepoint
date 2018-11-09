const spauth = require('node-sp-auth')
const axios = require('axios')

class Sharepoint {
  constructor (options) {
    this.url = options.url
    this.headers = null
    this.site = null
    this.formDigestValue = null
  }

  async authenticate (username, password) {
    const { headers } = await spauth.getAuth(this.url, { username, password })
    this.headers = {
      ...headers,
      Accept: 'application/json;odata=verbose'
    }
  }

  async getWebEndpoint () {
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

  async createFolder (dirPath, folderName) {
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

  async deleteFolder (dirPath, folderName) {
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

  async createFile (dirPath, fileName, data) {
    await axios({
      method: 'post',
      url: `${this.url}/_api/web/GetFolderByServerRelativeUrl('${this.site.serverRelativeUrl}${dirPath}')/Files/add(url='${fileName}')`,
      data,
      headers: {
        ...this.headers,
        'X-RequestDigest': this.formDigestValue
      }
    })
  }
}

module.exports = Sharepoint