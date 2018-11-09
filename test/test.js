/* eslint-env mocha */
'use strict'

const dotEnv = require('dotenv').config()
const chai = require('chai')
const expect = chai.expect

const Sharepoint = require('./../lib')

describe('Tests', function () {
  const FOLDER_NAME = 'TestFolder'
  const FILE_NAME = 'Test.txt'

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
    await sharepoint.createFolder(process.env.SHAREPOINT_DIR_PATH, FOLDER_NAME)
  })

  it('get directory contents, check new folder exists', async () => {
    const contents = await sharepoint.getContents(process.env.SHAREPOINT_DIR_PATH)
    expect(contents).to.not.eql(null)
    expect(contents.map(i => i.Name).includes(FOLDER_NAME)).to.eql(true)
  })

  it('get contents of new folder, should be empty', async () => {
    const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_DIR_PATH}/${FOLDER_NAME}`)
    expect(contents).to.eql([])
  })

  it('create file in new folder', async () => {
    await sharepoint.createFile(
      `${process.env.SHAREPOINT_DIR_PATH}/${FOLDER_NAME}`,
      FILE_NAME,
      'Testing 1 2 3...'
    )
  })

  it('get contents of new folder, expect new file', async () => {
    const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_DIR_PATH}/${FOLDER_NAME}`)
    expect(contents.length).to.eql(1)
    expect(contents[0].Name).to.eql('Test.txt')
  })

  it('delete the new file', async () => {
    await sharepoint.deleteFile(`${process.env.SHAREPOINT_DIR_PATH}/${FOLDER_NAME}`, FILE_NAME)
  })

  it('get contents of new folder, new file should be deleted', async () => {
    const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_DIR_PATH}/${FOLDER_NAME}`)
    expect(contents).to.eql([])
  })

  it('delete a folder', async () => {
    await sharepoint.deleteFolder(process.env.SHAREPOINT_DIR_PATH, FOLDER_NAME)
  })

  it('get directory contents, check folder has been deleted', async () => {
    const contents = await sharepoint.getContents(process.env.SHAREPOINT_DIR_PATH)
    expect(contents).to.not.eql(null)
    expect(contents.map(i => i.Name).includes(FOLDER_NAME)).to.eql(false)
  })
})