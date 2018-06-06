import createReport from './docxtemplate/indexNode';
import XlsxTemplate from 'xlsx-template'
import path from 'path'
import fse from "fs-extra"

function gendoc(templatePath, reportPath, mappingData, delimeter) {
  const ext = path.extname(templatePath)

  if (ext === '.docx') {
    createReport({
      template: templatePath,
      output: reportPath,
      data: mappingData,
      cmdDelimiter: '#',
      noSandbox: true // INSECURE - USE ONLY WITH TRUSTED TEMPLATES
    });
    return Promise.resolve()
  }
  else if (ext === '.xlsx') {
    return fse.readFile(templatePath)
      .then(data => {
        const xlsxTemplate = new XlsxTemplate(data)
        xlsxTemplate.substitute(1, mappingData)
        const report = xlsxTemplate.generate({ type: 'uint8array' })
        return fse.writeFile(reportPath, report)
      })
      .catch(e => {
        console.log(`something went wrong:`, e)
        return Promise.reject(e)
      })
  }
  else {
    throw new Error(`unsupported file extension ${ext}`)
  }
}

//#################
/*
// для тестирования шаблона нужно указывать имена файлов картинок которые ДОЛЖНЫ лежать рядом с запускаемым файлом (в эксплуатации используются id)
const DOCX_MAPPING_DATA = {
  name: 'John',
  lastname: 'Dow',
  images: [
    { id: 'card_1.png', label: 'ОПЦИОНАЛЬНОЕ ОПИСАНИЕ ФОТОЧКИ' },
    { id: 'card_2.png' },
    { id: 'card_3.png' },
    { id: 'card_4.png' }
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
    },
    {
      name: 'Lind Group',
      people: [
        {
          name: 'Reginald Lowe',
          projects: [
            {
              name: 'Reginald`s project 1'
            },
            {
              name: 'Reginald`s project 2'
            }
          ]
        },
        {
          name: 'Raul Tucker',
          projects: [
            {
              name: 'Raul`s project 1'
            },
            {
              name: 'Raul`s project 2'
            }
          ]
        }
      ]
    }
  ]
}

const PATH_TO_DOCX_TEMPLATE_FILE = path.resolve(path.dirname(require.main.filename), 'template.docx')

const PATH_TO_DOCX_REPORT_FILE = path.resolve(path.dirname(require.main.filename), 'report.docx')

const PLACEHOLDER_DELIMETER = '#' // default value - may be ommited

// ########################

const XLSX_MAPPING_DATA = {
  title: 'StarWars Persons',
  persons: [
    {
      firstName: 'Anakin',
      lastName: 'Skywalker'
    },
    {
      firstName: 'Lea',
      lastName: 'Organa'
    },
    {
      firstName: 'Han',
      lastName: 'Solo'
    },
  ]
}

const PATH_TO_XLSX_TEMPLATE_FILE = path.resolve(path.dirname(require.main.filename), 'template.xlsx')

const PATH_TO_XLSX_REPORT_FILE = path.resolve(path.dirname(require.main.filename), 'report.xlsx')

// gendoc(PATH_TO_DOCX_TEMPLATE_FILE, PATH_TO_DOCX_REPORT_FILE, DOCX_MAPPING_DATA, PLACEHOLDER_DELIMETER)
//   .then(() => {
//     console.log('docx generate complete')
//   })
//   .catch(e => {
//     console.error(e)
//   })


// gendoc(PATH_TO_XLSX_TEMPLATE_FILE, PATH_TO_XLSX_REPORT_FILE, XLSX_MAPPING_DATA)
//   .then(() => {
//     console.log('xlsx generate complete')
//   })
//   .catch(e => {
//     console.error(e)
//   })


//#################
*/
export default gendoc