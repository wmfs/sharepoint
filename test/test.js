/* eslint-env mocha */
'use strict'

const dotEnv = require('dotenv').config()
const chai = require('chai')
const expect = chai.expect

const Sharepoint = require('./../lib')

describe('Tests', function () {
  const NEW_FOLDER_NAME = 'TestFolder'

  let sharepoint

  before(function () {
    if (!(
      process.env.SHAREPOINT_URL &&
      process.env.SHAREPOINT_USERNAME &&
      process.env.SHAREPOINT_PASSWORD &&
      process.env.SHAREPOINT_DIR_PATH &&
      process.env.SHAREPOINT_SITE
    )) {
      console.log('Missing environment variables, skipping tests.')
      this.skip()
    }
  })

  it('create a new Sharepoint', () => {
    sharepoint = new Sharepoint({ url: process.env.SHAREPOINT_URL })
    expect(sharepoint.url).to.eql(process.env.SHAREPOINT_URL)
  })
  
  it('authenticate', async () => {
    await sharepoint.authenticate(process.env.SHAREPOINT_USERNAME, process.env.SHAREPOINT_PASSWORD)
    expect(sharepoint.headers.Cookie).to.not.eql(null)
    expect(sharepoint.headers.Accept).to.not.eql(null)
  })

  it('call the web endpoint', async () => {
    await sharepoint.getWebEndpoint()
    expect(sharepoint.site).to.not.eql(null)
    expect(sharepoint.site.id).to.not.eql(null)
    expect(sharepoint.site.description).to.not.eql(null)
    expect(sharepoint.site.created).to.not.eql(null)
    expect(sharepoint.site.serverRelativeUrl).to.not.eql(null)
    expect(sharepoint.site.lastModified).to.not.eql(null)
  })

  it('get form digest value', async () => {
    await sharepoint.getFormDigestValue()
    expect(sharepoint.formDigestValue).to.not.eql(null)
  })

  it('create a folder', async () => {
    await sharepoint.createFolder(process.env.SHAREPOINT_DIR_PATH, NEW_FOLDER_NAME)
  })

  it('get directory contents, check new folder exists', async () => {
    const contents = await sharepoint.getContents(process.env.SHAREPOINT_DIR_PATH)
    expect(contents).to.not.eql(null)
    expect(contents.map(i => i.Name).includes(NEW_FOLDER_NAME)).to.eql(true)
  })

  // delete new folder
})