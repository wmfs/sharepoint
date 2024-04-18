/* eslint-env mocha */
'use strict'

const chai = require('chai')
const expect = chai.expect
const Sharepoint = require('./../lib')
const process = require('node:process')
const path = require('node:path')
const fs = require('node:fs')

describe('Tests', function () {
  this.timeout(15000)

  const FOLDER_NAME = 'TestFolder'
  const FILE_NAME = 'Test.txt'
  const FILE_NAME_1 = 'Test.png'

  let sharepoint

  before(function () {
    if (!(
      process.env.SHAREPOINT_AUTH_SCOPE &&
      process.env.SHAREPOINT_CERT_FINGERPRINT &&
      process.env.SHAREPOINT_CERT_PASSPHRASE &&
      process.env.SHAREPOINT_CERT_PRIVATE_KEY_FILE &&
      process.env.SHAREPOINT_CLIENT_ID &&
      process.env.SHAREPOINT_TENANT_ID &&
      process.env.SHAREPOINT_URL &&
      process.env.SHAREPOINT_TESTS_DIR_PATH
    )) {
      console.log('Missing environment variables, skipping tests.')
      this.skip()
    }
  })

  it('construct Sharepoint instance', () => {
    sharepoint = new Sharepoint(process.env.SHAREPOINT_URL)
    expect(sharepoint.siteUrl).to.eql(process.env.SHAREPOINT_URL)
  })

  it('authenticate', async () => {
    await sharepoint.authenticate()
    expect(sharepoint.accessToken).to.not.eql(null)
  })

  it('call the web endpoint', async () => {
    await sharepoint.getWebEndpoint()
    expect(sharepoint.baseUrl).to.not.eql(null)
    expect(sharepoint.encodedBaseUrl).to.not.eql(null)
  })

  it('get form digest value', async () => {
    const formDigestValue = await sharepoint.getFormDigestValue()
    expect(formDigestValue).to.not.eql(null)
  })

  it('attempt to create a folder, without passing in a path', async () => {
    let error

    try {
      await sharepoint.createFolder()
    } catch (e) {
      error = e.message
    }

    expect(error).to.eql('You must provide a path.')
  })

  it('attempt to delete a folder, without passing in a path', async () => {
    let error

    try {
      await sharepoint.deleteFolder()
    } catch (e) {
      error = e.message
    }

    expect(error).to.eql('You must provide a path.')
  })

  it('attempt to create a file, without passing in a file name', async () => {
    let error

    try {
      await sharepoint.createFile({
        path: process.env.SHAREPOINT_TESTS_DIR_PATH,
        data: '...'
      })
    } catch (e) {
      error = e.message
    }

    expect(error).to.eql('You must provide a file name.')
  })

  it('attempt to create a file, without passing in data', async () => {
    let error

    try {
      await sharepoint.createFile({
        path: process.env.SHAREPOINT_TESTS_DIR_PATH,
        fileName: 'new file'
      })
    } catch (e) {
      error = e.message
    }

    expect(error).to.eql('You must provide data.')
  })

  it('attempt to delete a file, without passing in a file name', async () => {
    let error

    try {
      await sharepoint.deleteFile({
        path: process.env.SHAREPOINT_TESTS_DIR_PATH
      })
    } catch (e) {
      error = e.message
    }

    expect(error).to.eql('You must provide a file name.')
  })

  it('create a folder', async () => {
    await sharepoint.createFolder(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME}`)
  })

  it('get directory contents, check new folder exists', async () => {
    const contents = await sharepoint.getContents(process.env.SHAREPOINT_TESTS_DIR_PATH)
    expect(contents).to.not.eql(null)
    expect(contents.map(i => i.Name).includes(FOLDER_NAME)).to.eql(true)
  })

  it('get contents of new folder, should be empty', async () => {
    const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME}`)
    expect(contents).to.eql([])
  })

  it('create file in new folder', async () => {
    await sharepoint.createFile({
      path: `${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME}`,
      fileName: FILE_NAME,
      data: 'Testing 1 2 3...'
    })
  })

  it('get contents of new folder, expect new file', async () => {
    const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME}`)
    expect(contents.length).to.eql(1)
    expect(contents[0].Name).to.eql(FILE_NAME)
  })

  it('delete the new file', async () => {
    await sharepoint.deleteFile({
      path: `${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME}`,
      fileName: FILE_NAME
    })
  })

  it('get contents of new folder, new file should be deleted', async () => {
    const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME}`)
    expect(contents).to.eql([])
  })

  it('upload file read in from fixtures', async () => {
    const data = getBinaryData(path.resolve(__dirname, 'fixtures', FILE_NAME))

    await sharepoint.createFile({
      path: `${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME}`,
      fileName: FILE_NAME,
      data
    })
  })

  it('get contents of new folder, expect new file from fixtures', async () => {
    const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME}`)
    expect(contents.length).to.eql(1)
    expect(contents[0].Name).to.eql(FILE_NAME)
  })

  it('upload file of different format (png) from fixtures', async () => {
    const data = getBinaryData(path.resolve(__dirname, 'fixtures', FILE_NAME_1))

    await sharepoint.createFile({
      path: `${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME}`,
      fileName: FILE_NAME_1,
      data
    })
  })

  it('get contents of new folder, expect new file of different format (png) from fixtures', async () => {
    const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME}`)
    expect(contents.length).to.eql(2)
    expect(contents.map(i => i.Name).includes(FILE_NAME_1)).to.eql(true)
  })

  it('upload file read in from fixtures using chunks', async () => {
    const filePath = path.resolve(__dirname, 'fixtures', FILE_NAME_1)
    const { size } = fs.statSync(filePath)
    const stream = fs.createReadStream(filePath, { highWaterMark: 1024 * 2 })
    await sharepoint.createFileChunked({
      path: `${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME}`,
      fileName: FILE_NAME_1,
      stream,
      fileSize: size,
      chunkSize: 1024 * 2
    })
  })

  it('delete a folder', async () => {
    await sharepoint.deleteFolder(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME}`)
  })

  it('get directory contents, check folder has been deleted', async () => {
    const contents = await sharepoint.getContents(process.env.SHAREPOINT_TESTS_DIR_PATH)
    expect(contents).to.not.eql(null)
    expect(contents.map(i => i.Name).includes(FOLDER_NAME)).to.eql(false)
  })
})

function getBinaryData (filepath) {
  const base64 = fs.readFileSync(filepath, { encoding: 'base64' })
  const encodedBase64String = base64.replace(/^data:+[a-z]+\/+[a-z]+;base64,/, '')
  return Buffer.from(encodedBase64String, 'base64')
}
