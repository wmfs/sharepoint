/* eslint-env mocha */
'use strict'

const dotEnv = require('dotenv').config()
const path = require('path')
const fs = require('fs')
const chai = require('chai')
const expect = chai.expect

const Sharepoint = require('./../lib')

describe('Tests', function () {
  this.timeout(15000)

  const FOLDER_NAME = 'TestFolder'
  const FILE_NAME = 'Test.txt'
  const FILE_NAME_1 = 'Test.png'

  let sharepoint

  before(function () {
    if (!(
      process.env.SHAREPOINT_URL &&
      process.env.SHAREPOINT_USERNAME &&
      process.env.SHAREPOINT_PASSWORD &&
      process.env.SHAREPOINT_DIR_PATH
    )) {
      console.log('Missing environment variables, skipping tests.')
      this.skip()
    }
  })

  it('attempt to create a new Sharepoint without passing url', () => {
    let error

    try {
      const sp = new Sharepoint()
    } catch (e) {
      error = e.message
    }

    expect(error).to.eql('You must provide a url.')
  })

  it('create a new Sharepoint', () => {
    sharepoint = new Sharepoint(process.env.SHAREPOINT_URL)
    expect(sharepoint.url).to.eql(process.env.SHAREPOINT_URL)
  })

  it('attempt to authenticate without passing in username or password', async () => {
    let error
    
    try {
      await sharepoint.authenticate()
    } catch (e) {
      error = e.message
    }

    expect(error).to.eql('You must provide a username and password.')
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

  it('attempt to create a folder, without passing in a folder name', async () => {
    let error

    try {
      await sharepoint.createFolder({
        dirPath: process.env.SHAREPOINT_DIR_PATH
      })
    } catch (e) {
      error = e.message
    }
    
    expect(error).to.eql('You must provide a folder name.')
  })

  it('create a folder', async () => {
    await sharepoint.createFolder({
      dirPath: process.env.SHAREPOINT_DIR_PATH,
      folderName: FOLDER_NAME
    })
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
    await sharepoint.createFile({
      dirPath: `${process.env.SHAREPOINT_DIR_PATH}/${FOLDER_NAME}`,
      fileName: FILE_NAME,
      data: 'Testing 1 2 3...'
    })
  })

  it('get contents of new folder, expect new file', async () => {
    const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_DIR_PATH}/${FOLDER_NAME}`)
    expect(contents.length).to.eql(1)
    expect(contents[0].Name).to.eql(FILE_NAME)
  })

  it('delete the new file', async () => {
    await sharepoint.deleteFile({
      dirPath: `${process.env.SHAREPOINT_DIR_PATH}/${FOLDER_NAME}`,
      fileName: FILE_NAME
    })
  })

  it('get contents of new folder, new file should be deleted', async () => {
    const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_DIR_PATH}/${FOLDER_NAME}`)
    expect(contents).to.eql([])
  })

  it('upload file read in from fixtures', async () => {
    const data = getBinaryData(path.resolve(__dirname, 'fixtures', FILE_NAME))

    await sharepoint.createFile({
      dirPath: `${process.env.SHAREPOINT_DIR_PATH}/${FOLDER_NAME}`,
      fileName: FILE_NAME,
      data
    })    
  })

  it('get contents of new folder, expect new file from fixtures', async () => {
    const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_DIR_PATH}/${FOLDER_NAME}`)
    expect(contents.length).to.eql(1)
    expect(contents[0].Name).to.eql(FILE_NAME)
  })

  it('upload file of different format (png) from fixtures', async () => {
    const data = getBinaryData(path.resolve(__dirname, 'fixtures', FILE_NAME_1))

    await sharepoint.createFile({
      dirPath: `${process.env.SHAREPOINT_DIR_PATH}/${FOLDER_NAME}`,
      fileName: FILE_NAME_1,
      data
    })
  })

  it('get contents of new folder, expect new file of different format (png) from fixtures', async () => {
    const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_DIR_PATH}/${FOLDER_NAME}`)
    expect(contents.length).to.eql(2)
    expect(contents.map(i => i.Name).includes(FILE_NAME_1)).to.eql(true)
  })

  it('delete a folder', async () => {
    await sharepoint.deleteFolder({
      dirPath: process.env.SHAREPOINT_DIR_PATH,
      folderName: FOLDER_NAME
    })
  })

  it('get directory contents, check folder has been deleted', async () => {
    const contents = await sharepoint.getContents(process.env.SHAREPOINT_DIR_PATH)
    expect(contents).to.not.eql(null)
    expect(contents.map(i => i.Name).includes(FOLDER_NAME)).to.eql(false)
  })
})

function getBinaryData (filepath) {
  const base64 = fs.readFileSync(filepath, { encoding: 'base64' })
  const encodedBase64String = base64.replace(/^data:+[a-z]+\/+[a-z]+;base64,/, '')
  return Buffer.from(encodedBase64String, 'base64')
}