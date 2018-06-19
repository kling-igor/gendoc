GenDoc
======

A small utility to genereate .docx and .xlsx by specified template and mapping data.

## Requirements

Dependent on *babel-runtime*

## Installation

Via npm:

```sh
  npm install @igorkling/gendoc -S
```  

Or using yarn:

```sh
  yarn add @igorkling/gendoc
```  

## Usage

### Generating .xlsx document

```js
const gendoc = require('@igorkling/gendoc').default
const path = require('path')

// assume template and result are one the same lavel as binary
const PATH_TO_XLSX_TEMPLATE_FILE = path.resolve(path.dirname(require.main.filename), 'template.xlsx')
const PATH_TO_XLSX_REPORT_FILE = path.resolve(path.dirname(require.main.filename), 'report.xlsx')


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

const generator = new GenDoc()

generator.createReport(PATH_TO_XLSX_TEMPLATE_FILE, PATH_TO_XLSX_REPORT_FILE, XLSX_MAPPING_DATA)
  .then(() => {
    console.log('xlsx generate complete')
  })
  .catch(e => {
    console.error(e)
  })
```

### Generating .docx document

```js
const DOCX_MAPPING_DATA = {
  name: 'John',
  lastname: 'Dow',
  images: [
    { path: 'image_1.png', label: 'Some optional image description' },
    { path: 'image_2.png' },
    { path: 'image_3.png' },
    { path: 'image_4.png' }
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

const PATH_TO_DOCX_TEMPLATE_FILE = path.resolve(path.dirname(require.main.filename), 'template.docx')

const PATH_TO_DOCX_REPORT_FILE = path.resolve(path.dirname(require.main.filename), 'report.docx')

const PLACEHOLDER_DELIMETER = '#' // default value - may be ommited

const readFile = async (imageInfo) => {
  // path is the module name
  const { path: filename } = imageInfo

  const imagePath = path.join(processPath, filename)
  const extension = path.extname(filename).substring(1)

  const buffer = await fse.readFile(imagePath)

  const base64 = buffer.toString('base64')

  return { extension, data: base64 }
}

const generator = new GenDoc(readFile)

generator.createReport(PATH_TO_DOCX_TEMPLATE_FILE, PATH_TO_DOCX_REPORT_FILE, XLSX_MAPPING_DATA, PLACEHOLDER_DELIMETER)
  .then(() => {
    console.log('docx generate complete')
  })
  .catch(e => {
    console.error(e)
  })
```

It is up to you how specify images in mapping data - as local filenames or web links.
You have to provide image reading function. If there are no images expected in .docx then
image reading function can be ommited.


## Release History

* 2.0.0 Image reading function
* 1.0.0 Initial release
