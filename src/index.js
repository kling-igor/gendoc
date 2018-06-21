import XlsxTemplate from 'xlsx-template'
// import { decode } from 'base64-arraybuffer'
import path from 'path'
import fse from "fs-extra"
import DocxTemplate from './docxtemplate'


const readFile = async (imageInfo) => {
  throw new Error('you should pass function to GenDoc to handle IMAGE tags data in docx template')
}


class GenDoc {

  // { extension: ext, data: base64Buffer }

  /**
   *Creates an instance of GenDoc.
   * @param {Interface} fileService - an interface with readFile method
   * @memberof GenDoc
   */
  constructor(onReadImage = readFile) {
    this.docxTemplate = new DocxTemplate(onReadImage)
  }

  /**
   * create report asynchronously
   * @param {string} templatePath - template file path
   * @param {string} reportPath - report file path
   * @param {object} mappingData - data to be visualized in the report
   * @param {string} delimeter - character or characters sequence in template
   * @return {Promise} - a promise resolved by write file operation finish
   */
  async createReport(templatePath, reportPath, mappingData, delimeter = '#') {
    const ext = path.extname(templatePath)

    if (ext === '.docx') {
      return this.docxTemplate.createReport({
        template: templatePath,
        reportPath,
        data: mappingData,
        cmdDelimiter: delimeter,
        noSandbox: true // INSECURE - USE ONLY WITH TRUSTED TEMPLATES
      })
    } else if (ext === '.xlsx') {
      return fse.readFile(templatePath)
        .then(data => {
          const xlsxTemplate = new XlsxTemplate(data)
          xlsxTemplate.substitute(1, mappingData)
          const report = xlsxTemplate.generate({ type: 'uint8array' })
          return fse.writeFile(reportPath, report)
        })
        .catch(e => {
          console.error(`something went wrong:`, e)
          return Promise.reject(e)
        })
    }
    throw new Error(`unsupported file extension ${ext}`)
  }
}

export default GenDoc



// HOW TO USE!!!
/*

const processPath = path.dirname(require.main.filename)

const readFile = async (imageInfo) => {
  // path is the module name
  const { path: filename } = imageInfo

  const imagePath = path.join(processPath, filename)
  const extension = path.extname(filename).substring(1)

  const buffer = await fse.readFile(imagePath)

  const base64 = buffer.toString('base64')

  return { extension, data: base64 }
}

const genDoc = new GenDoc(readFile)

const DOCX_MAPPING_DATA = {
  name: 'John',
  lastname: 'Dow',
  images: [
    { id: 'image_1.png', label: 'Some optional image description' },
    { id: 'image_2.png' },
    { id: 'image_3.png' },
    { id: 'image_4.png' }
  ],
  companies: [
    {
      name: 'Treutel Inc',
      people: [
        {
          name: 'Romain Hoogmoed',
          projects: [
            {
              name: 'Romain`s project 1'
            },
            {
              name: 'Romain`s project 2'
            }
          ]
        },
        {
          name: 'Leo Fuller',
          projects: [
            {
              name: 'Leo`s project 1'
            },
            {
              name: 'Leo`s project 2'
            }
          ]
        }
      ]
    },
    {
      name: 'Mante and Sons',
      people: [
        {
          name: 'Rosa Powell',
          projects: [
            {
              name: 'Rosa`s project 1'
            },
            {
              name: 'Rosa`s project 2'
            }
          ]
        },
        {
          name: 'Carter Moreno',
          projects: [
            {
              name: 'Carter`s project 1'
            },
            {
              name: 'Carter`s project 2'
            }
          ]
        }
      ]
    }
  ]
}

const templatePath = path.join(processPath, 'template.docx')
const reportPath = path.join(processPath, 'report.docx')

genDoc.createReport(templatePath, reportPath, DOCX_MAPPING_DATA, '#')
  .then(() => {
    console.log('DONE')
  })
  .catch(e => {
    console.error(e)
  })

*/