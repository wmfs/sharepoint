const spauth = require('node-sp-auth')
const dotEnv = require('dotenv')
const axios = require('axios')
dotEnv.config() // Allows use of .env file for setting env variables, if preferred.

module.exports = async function () {
  console.log('SharePoint Authentication')
  console.log('-------------------------')

  // Grab URL from env variables
  const url = process.env.SHAREPOINT_URL
  console.log(`SharePoint URL: '${url}'`)

  const rootSharePointDirPath = process.env.SHAREPOINT_DIR_PATH
  console.log(`rootSharePointDirPath: '${rootSharePointDirPath}'`)

  // Grab credentials from env variables
  const credentials = {
    username: process.env.SHAREPOINT_USERNAME,
    password: process.env.SHAREPOINT_PASSWORD
  }
  console.log(`Username: ${credentials.username}`)

  try {
    const {headers} = await spauth.getAuth(url, credentials)

    console.log(`Got cookie? ${headers.hasOwnProperty('Cookie')}`)
    console.log(`Auth type: ${headers.Cookie.split('=')[0]}`)

    headers['Accept'] = 'application/json;odata=verbose'

    // Call the web endpoint
    const {data} = await axios.get(`${url}/_api/web`, {headers, responseType: 'json'})

    const site = data.d
    const ServerRelativeUrl = site.ServerRelativeUrl
    console.log('')
    console.log('Site details')
    console.log('------------')
    console.log(`ID: ${site.Id}`)
    console.log(`Title: ${site.Title}`)
    console.log(`Description: ${site.Description}`)
    console.log(`Created: ${site.Created}`)
    console.log(`ServerRelativeUrl: ${ServerRelativeUrl}`)
    console.log(`Modified: ${site.LastItemUserModifiedDate}`)
    console.log('')

    // Get a "form digest value" (time limited)
    const formDigestValue = await getFormDigestValue(url, headers)
    console.log(`formDigestValue: ${formDigestValue}`)
    console.log('')

    // Create a new folder
    const createFolderResponse = await createAFolder('Hello', url, formDigestValue, headers)
    console.log('')
    console.log(`Created folder './${createFolderResponse.Name}' at ${createFolderResponse.TimeCreated} (UniqueId=${createFolderResponse.UniqueId})`)
    console.log('')

    // Get file contents
    const contents = await getContents(url, headers)
    console.log('Contents:\n', contents.map(i => i.Name).join('\n'))
  } catch (e) {
    if (e.response) {
      console.log(`${e.response.status}: ${e.response.statusText}`)
      console.log(JSON.stringify(e.response.data))
    } else console.error(e)
  }
}

const getFormDigestValue = async (url, headers) => {
  const contextInfo = await axios({
    method: 'post',
    url: `${url}/_api/contextinfo`,
    headers: {
      ...headers,
      'content-type': 'application/json;odata=verbose'
    },
    responseType: 'json'
  })
  return contextInfo.data.d.GetContextWebInformation.FormDigestValue
}

const getContents = async (url, headers) => {
  const get = type => {
    return axios.get(
      `${url}/_api/web/GetFolderByServerRelativeUrl('${process.env.SHAREPOINT_DIR_PATH}')/${type}`,
      {headers, responseType: 'json'}
    )
  }

  const folders = await get('Folders')
  const files = await get('Files')

  return [...folders.data.d.results, ...files.data.d.results]
}

// https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest
const createAFolder = async (folderName, url, formDigestValue, headers) => {
  const result = await axios({
    method: 'post',
    url: `${url}/_api/web/folders`,
    headers: {
      ...headers,
      'content-type': 'application/json;odata=verbose',
      'X-RequestDigest': formDigestValue
    },
    data: {
      __metadata: {type: 'SP.Folder' },
      ServerRelativeUrl: `${process.env.SHAREPOINT_DIR_PATH}/${folderName}`
    },
    responseType: 'json'
  })
  return result.data.d
}
