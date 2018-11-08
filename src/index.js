const spauth = require('node-sp-auth')
const dotEnv = require('dotenv')
const axios = require('axios')
dotEnv.config() // Allows use of .env file for setting env variables, if preferred.

const SECURITY_TOKEN_REQUEST_ENVELOPE = '<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\" xmlns:a=\"http://www.w3.org/2005/08/addressing\">' + 
'<s:Header><a:Action s:mustUnderstand=\"1\">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action>' + 
'<a:ReplyTo><a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address></a:ReplyTo>' + 
'<a:To s:mustUnderstand=\"1\">https://login.microsoftonline.com/extSTS.srf</a:To>' + 
'<o:Security s:mustUnderstand=\"1\" xmlns:o=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd\">' + 
'<o:UsernameToken><o:Username>[USERNAME]</o:Username><o:Password>[PASSWORD]</o:Password></o:UsernameToken>' + 
'</o:Security></s:Header><s:Body><t:RequestSecurityToken xmlns:t=\"http://schemas.xmlsoap.org/ws/2005/02/trust\">' + 
'<wsp:AppliesTo xmlns:wsp=\"http://schemas.xmlsoap.org/ws/2004/09/policy\"><a:EndpointReference>' + 
'<a:Address>[SITE]</a:Address></a:EndpointReference></wsp:AppliesTo>' + 
'<t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType>' + 
'<t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType>' + 
'<t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType>' + 
'</t:RequestSecurityToken></s:Body></s:Envelope>'

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

    const securityToken = SECURITY_TOKEN_REQUEST_ENVELOPE
      .replace('[USERNAME]', credentials.username)
      .replace('[PASSWORD]', credentials.password)
      .replace('[SITE]', process.env.SHAREPOINT_SITE)

    // get access token
    const securityTokenRes = await axios.post(
      process.env.SECURITY_TOKEN_URL,
      securityToken,
      { headers, responseType: 'json' }
    )

    console.log('>>', securityTokenRes.data)

    // Get context info
    // const contextInfo = await axios.post(`${url}/_api/contextinfo`, { headers, responseType: 'json' })
    // console.log('>>>', contextInfo)

    // Get the contents
    // const contents = await getContents(url, headers)
    // console.log('Contents:', contents.map(i => i.Name).join(', '))
    // console.log('')

    // Attempt to create a folder
    // const createFolderRes = await axios({
    //   method: 'post',
    //   url: `${url}/_api/web/Folders`,
    //   header: {
    //     ...headers,
    //     'X-RequestDigest': 'form digest value',
    //     'content-type': 'application/json;odata=verbose'
    //   },
    //   responseType: 'json',
    //   data: {
    //     '__metadata': { 'type': 'SP.Folder' },
    //     'ServerRelativeUrl': '/Shared Documents/General/TYMLYNODE/NewTestFolder'
    //   }
    // })

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
