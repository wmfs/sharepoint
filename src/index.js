const spauth = require('node-sp-auth')
const dotEnv = require('dotenv')
const axios = require('axios')
dotEnv.config() // Allows use of .env file for setting env variables, if preferred.

module.exports = async function () {
  console.log('SharePoint Authentication')
  console.log('-------------------------')

  // Grab URL from env variables
  const url = process.env.SHAREPOINT_URL
  console.log(`SharePoint URL: ${url}`)

  // Grab credentials from env variables
  const credentials = {
    username: process.env.SHAREPOINT_USERNAME,
    password: process.env.SHAREPOINT_PASSWORD
  }
  console.log(`Username: ${credentials.username}`)
  
  try {
    const { headers } = await spauth.getAuth(url, credentials)

    console.log(`Got cookie? ${headers.hasOwnProperty('Cookie')}`)
    console.log(`Auth type: ${headers.Cookie.split('=')[0]}`)
    
    headers['Accept'] = 'application/json;odata=verbose'

    // Call the web endpoint
    const { data } = await axios.get(`${url}/_api/web`, { headers, responseType: 'json' })

    const site = data.d
    
    console.log('')
    console.log('Site details')
    console.log('------------')
    console.log(`ID: ${site.Id}`)
    console.log(`Title: ${site.Title}`)
    console.log(`Description: ${site.Description}`)
    console.log(`Created: ${site.Created}`)
    console.log(`Modified: ${site.LastItemUserModifiedDate}`)
    console.log('')

    // Get the contents
    const contents = await getContents(url, headers)
    console.log('Contents:', contents.map(i => i.Name).join(', '))

  } catch (e) {
    if (e.response) {
      console.error(`${e.response.status}: ${e.response.statusText}`)
      console.error(e.response.data.error.code)
      console.error(e.response.data.error.message.value)
    } else console.error(e)
  }
}

const getContents = async (url, headers) => {
  const get = type => {
    return axios.get(
      `${url}/_api/web/GetFolderByServerRelativeUrl('${process.env.SHAREPOINT_DIR_PATH}')/${type}`,
      { headers, responseType: 'json' }
    )
  }

  const folders = await get('Folders')
  const files = await get('Files')

  return [...folders.data.d.results, ...files.data.d.results]
}
