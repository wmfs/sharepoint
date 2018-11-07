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

    // Attempt to create a folder

    // const createFolderRes = await axios.post(
    //   `${url}/_api/web/folders/add('Documents/TEST_FOLDER')`,
    //   { 
    //     headers: {
    //       ...headers,
    //       // 'X-RequestDigest': 'form digest value', // <--- dunno if this lot are necessary?
    //       // 'content-type': 'application/json;odata=verbose',
    //       // 'content-length': 'length of post body'
    //     },
    //     responseType: 'json'
    //   }
    // )

    const createFolderRes = await axios({
      method: 'post',
      url: `${url}/_api/web/folders`,
      header: {
        ...headers,
        'X-RequestDigest': 'form digest value',
        'content-type': 'application/json;odata=verbose',
        'content-length': 'length of post body'
      },
      responseType: 'json',
      data: {
        '__metadata': { 'type': 'SP.Folder' },
        'ServerRelativeUrl': '/Documents/General/TYMLYNODE/NewTestFolder'
      }
    })

  } catch (e) {
    if (e.response) {
      console.log(`${e.response.status}: ${e.response.statusText}`) 
      console.log(JSON.stringify(e.response.data))
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
