/* eslint-env mocha */
'use strict'

const chai = require('chai')
const expect = chai.expect
const Sharepoint = require('./../lib')
const process = require('node:process')
const path = require('node:path')
const fs = require('node:fs')

describe('tests', function () {
  this.timeout(15000)

  const FOLDER_NAME1 = 'TestFolder1'
  const FOLDER_NAME2 = 'TestFolder8834634'
  const FOLDER_NAME3 = 'TestFolder0'
  const FOLDER_NAME4 = 'TestFolder396'

  const FILE_NAME1 = 'test-file-453446.txt'
  const FILE_NAME2 = 'test-file-5.txt'
  const FILE_NAME3 = 'test-file-0.txt'
  const FILE_NAME4 = 'test-file-10.txt'

  const TEXT_FILE_FILENAME = 'Test.txt'
  const BINARY_FILE_FILENAME = 'Test.png'

  const authScope = process.env.SHAREPOINT_AUTH_SCOPE
  const certFingerprint = process.env.SHAREPOINT_CERT_FINGERPRINT
  const certPassphrase = process.env.SHAREPOINT_CERT_PASSPHRASE
  const certPrivateKeyFile = process.env.SHAREPOINT_CERT_PRIVATE_KEY_FILE
  const clientId = process.env.SHAREPOINT_CLIENT_ID
  const tenantId = process.env.SHAREPOINT_TENANT_ID

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

  describe('check for handling of missing environment variables', () => {
    it('construct Sharepoint instance with no auth scope env var set', () => {
      delete process.env.SHAREPOINT_AUTH_SCOPE
      try {
        new Sharepoint(process.env.SHAREPOINT_URL) // eslint-disable-line
      } catch (err) {
        expect(err.message).to.eql('SHAREPOINT_AUTH_SCOPE environment variable has not been set')
      }
    })

    it('construct Sharepoint instance with auth scope env var set to an invalid value', () => {
      process.env.SHAREPOINT_AUTH_SCOPE = 'https://invalid-value'
      try {
        new Sharepoint(process.env.SHAREPOINT_URL) // eslint-disable-line
      } catch (err) {
        expect(err.message).to.eql('SHAREPOINT_AUTH_SCOPE environment variable value is not valid - it must begin with "https://" and end with ".sharepoint.com/.default"')
      }
    })

    it('construct Sharepoint instance with no client id env var set', () => {
      process.env.SHAREPOINT_AUTH_SCOPE = authScope
      delete process.env.SHAREPOINT_CLIENT_ID
      try {
        new Sharepoint(process.env.SHAREPOINT_URL) // eslint-disable-line
      } catch (err) {
        expect(err.message).to.eql('SHAREPOINT_CLIENT_ID environment variable has not been set')
      }
    })

    it('construct Sharepoint instance with no tenant id env var set', () => {
      process.env.SHAREPOINT_CLIENT_ID = clientId
      delete process.env.SHAREPOINT_TENANT_ID
      try {
        new Sharepoint(process.env.SHAREPOINT_URL) // eslint-disable-line
      } catch (err) {
        expect(err.message).to.eql('SHAREPOINT_TENANT_ID environment variable has not been set')
      }
    })

    it('construct Sharepoint instance with no cert passphrase env var set', () => {
      process.env.SHAREPOINT_TENANT_ID = tenantId
      delete process.env.SHAREPOINT_CERT_PASSPHRASE
      try {
        new Sharepoint(process.env.SHAREPOINT_URL) // eslint-disable-line
      } catch (err) {
        expect(err.message).to.eql('SHAREPOINT_CERT_PASSPHRASE environment variable has not been set')
      }
    })

    it('construct Sharepoint instance with no cert fingerprint env var set', () => {
      process.env.SHAREPOINT_CERT_PASSPHRASE = certPassphrase
      delete process.env.SHAREPOINT_CERT_FINGERPRINT
      try {
        new Sharepoint(process.env.SHAREPOINT_URL) // eslint-disable-line
      } catch (err) {
        expect(err.message).to.eql('SHAREPOINT_CERT_FINGERPRINT environment variable has not been set')
      }
    })

    it('construct Sharepoint instance with cert fingerprint env var set to an invalid value', () => {
      process.env.SHAREPOINT_CERT_FINGERPRINT = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
      try {
        new Sharepoint(process.env.SHAREPOINT_URL) // eslint-disable-line
      } catch (err) {
        expect(err.message).to.eql('SHAREPOINT_CERT_FINGERPRINT environment variable value is not valid - it must be exactly 40 characters in length')
      }
    })

    it('construct Sharepoint instance with no cert private key file env var set', () => {
      process.env.SHAREPOINT_CERT_FINGERPRINT = certFingerprint
      delete process.env.SHAREPOINT_CERT_PRIVATE_KEY_FILE
      try {
        new Sharepoint(process.env.SHAREPOINT_URL) // eslint-disable-line
      } catch (err) {
        expect(err.message).to.eql('SHAREPOINT_CERT_PRIVATE_KEY_FILE environment variable has not been set')
      }
    })

    it('construct Sharepoint instance with no cert private key file env var set', () => {
      process.env.SHAREPOINT_CERT_PRIVATE_KEY_FILE = './fixtures/no-such-file.pem'
      try {
        new Sharepoint(process.env.SHAREPOINT_URL) // eslint-disable-line
      } catch (err) {
        expect(err.message).to.eql('specified sharepoint certificate private key file (\'./fixtures/no-such-file.pem\') does not exist')
      }
    })
  })

  describe('sharepoint interactivity tests', () => {
    it('attempt to construct Sharepoint instance without a siteUrl', () => {
      try {
        new Sharepoint() // eslint-disable-line
      } catch (err) {
        expect(err.message).to.eql('siteUrl has not been specified')
      }
    })

    it('construct Sharepoint instance', () => {
      process.env.SHAREPOINT_CERT_PRIVATE_KEY_FILE = certPrivateKeyFile
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

    it('get test directory contents, should be empty', async () => {
      const contents = await sharepoint.getContents(process.env.SHAREPOINT_TESTS_DIR_PATH)
      expect(contents.length).to.eql(0)
    })

    it('create main test folder', async () => {
      await sharepoint.createFolder(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME1}`)
    })

    it('create test folder 2', async () => {
      await sharepoint.createFolder(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME2}`)
    })

    it('create test folder 3', async () => {
      await sharepoint.createFolder(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME3}`)
    })

    it('create test folder 4', async () => {
      await sharepoint.createFolder(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME4}`)
    })

    it('get directory contents, check folders returned in expected natural order', async () => {
      const contents = await sharepoint.getContents(process.env.SHAREPOINT_TESTS_DIR_PATH)
      expect(contents.length).to.eql(4)
      expect(contents[0].Name).to.eql(FOLDER_NAME3)
      expect(contents[1].Name).to.eql(FOLDER_NAME1)
      expect(contents[2].Name).to.eql(FOLDER_NAME4)
      expect(contents[3].Name).to.eql(FOLDER_NAME2)
    })

    it('get contents of new folder, should be empty', async () => {
      const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME1}`)
      expect(contents).to.eql([])
    })

    it('create file in new folder', async () => {
      await sharepoint.createFile({
        path: `${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME1}`,
        fileName: TEXT_FILE_FILENAME,
        data: 'Testing 1 2 3...'
      })
    })

    it('get contents of new folder, expect new file', async () => {
      const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME1}`)
      expect(contents.length).to.eql(1)
      expect(contents[0].Name).to.eql(TEXT_FILE_FILENAME)
    })

    it('delete the new file', async () => {
      await sharepoint.deleteFile({
        path: `${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME1}`,
        fileName: TEXT_FILE_FILENAME
      })
    })

    it('get contents of new folder, new file should be deleted', async () => {
      const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME1}`)
      expect(contents).to.eql([])
    })

    it('upload text file with filename 1', async () => {
      const data = getBinaryData(path.resolve(__dirname, 'fixtures', TEXT_FILE_FILENAME))

      await sharepoint.createFile({
        path: `${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME1}`,
        fileName: FILE_NAME1,
        data
      })
    })

    it('upload text file with filename 2', async () => {
      const data = getBinaryData(path.resolve(__dirname, 'fixtures', TEXT_FILE_FILENAME))

      await sharepoint.createFile({
        path: `${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME1}`,
        fileName: FILE_NAME2,
        data
      })
    })

    it('upload text file with filename 3', async () => {
      const data = getBinaryData(path.resolve(__dirname, 'fixtures', TEXT_FILE_FILENAME))

      await sharepoint.createFile({
        path: `${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME1}`,
        fileName: FILE_NAME3,
        data
      })
    })

    it('upload text file with filename 4', async () => {
      const data = getBinaryData(path.resolve(__dirname, 'fixtures', TEXT_FILE_FILENAME))

      await sharepoint.createFile({
        path: `${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME1}`,
        fileName: FILE_NAME4,
        data
      })
    })

    it('get directory contents, check files returned in expected natural order', async () => {
      const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME1}`)
      expect(contents.length).to.eql(4)
      expect(contents[0].Name).to.eql(FILE_NAME3)
      expect(contents[1].Name).to.eql(FILE_NAME2)
      expect(contents[2].Name).to.eql(FILE_NAME4)
      expect(contents[3].Name).to.eql(FILE_NAME1)
    })

    it('upload file of different format (png) from fixtures', async () => {
      const data = getBinaryData(path.resolve(__dirname, 'fixtures', BINARY_FILE_FILENAME))

      await sharepoint.createFile({
        path: `${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME1}`,
        fileName: BINARY_FILE_FILENAME,
        data
      })
    })

    it('get contents of new folder, expect new png file from fixtures', async () => {
      const contents = await sharepoint.getContents(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME1}`)
      expect(contents.length).to.eql(5)
      expect(contents[4].Name).to.eql(BINARY_FILE_FILENAME)
    })

    it('upload file read in from fixtures using chunks', async () => {
      const filePath = path.resolve(__dirname, 'fixtures', BINARY_FILE_FILENAME)
      const { size } = fs.statSync(filePath)
      const stream = fs.createReadStream(filePath, { highWaterMark: 1024 * 2 })
      await sharepoint.createFileChunked({
        path: `${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME1}`,
        fileName: BINARY_FILE_FILENAME,
        stream,
        fileSize: size,
        chunkSize: 1024 * 2
      })
    })

    it('delete folder 1', async () => {
      await sharepoint.deleteFolder(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME1}`)
    })

    it('delete folder 2', async () => {
      await sharepoint.deleteFolder(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME2}`)
    })

    it('delete folder 3', async () => {
      await sharepoint.deleteFolder(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME3}`)
    })

    it('delete folder 4', async () => {
      await sharepoint.deleteFolder(`${process.env.SHAREPOINT_TESTS_DIR_PATH}/${FOLDER_NAME4}`)
    })

    it('get directory contents, check folders have been deleted', async () => {
      const contents = await sharepoint.getContents(process.env.SHAREPOINT_TESTS_DIR_PATH)
      expect(contents).to.not.eql(null)
      expect(contents.map(i => i.Name).includes(FOLDER_NAME1)).to.eql(false)
      expect(contents.map(i => i.Name).includes(FOLDER_NAME2)).to.eql(false)
      expect(contents.map(i => i.Name).includes(FOLDER_NAME3)).to.eql(false)
      expect(contents.map(i => i.Name).includes(FOLDER_NAME4)).to.eql(false)
    })
  })
})

function getBinaryData (filepath) {
  const base64 = fs.readFileSync(filepath, { encoding: 'base64' })
  const encodedBase64String = base64.replace(/^data:+[a-z]+\/+[a-z]+;base64,/, '')
  return Buffer.from(encodedBase64String, 'base64')
}
