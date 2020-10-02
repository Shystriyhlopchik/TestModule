import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import Docxtemplater from 'docxtemplater';
// @ts-ignore
import JSZip from 'jszip';
import JSZipUtils from 'jszip-utils';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'TestModule';
  fileName = 'ExcelSheet.xlsx';

  exportexcel(): void
  {
    // сюда передается идентификатор таблицы
    let element = document.getElementById ('excel-table');
    const ws: XLSX.WorkSheet = XLSX.utils.table_to_sheet (element);

    // создать книгу и добавить рабочий лист
    const wb: XLSX.WorkBook = XLSX.utils.book_new ();
    XLSX.utils.book_append_sheet (wb, ws, 'Sheet1');
    // console.log(wb.Sheets.Sheet1['!ref'] = 'A2:')
    // сохранить в файл
    // XLSX.writeFile (wb, this.fileName);
  }

  // loadFile(url, callback) {
  //   JSZipUtils.getBinaryContent(url, callback);
  // }
  //
  // test(val) {
  //   let form = document.forms;
  //   console.log(form);
  // }

  dataGeneration() {
    console.log()
  }

  exportDocx(files: Blob, centersStateTableData: object): void {
    const file = files[0]
    const reader = new FileReader()
    reader.readAsArrayBuffer(file)
    console.log(reader)
    // генерируем JSON
    // const data = this.createObjectData(this.listOfColumns, centersStateTableData)
    reader.onload = function() {
      const content = reader.result
      const zip = new JSZip(content)
      const doc = new Docxtemplater().loadZip(zip)
      const data = new Data(centersStateTableData)
      generatingJsonObject(data)
      doc.setData(data)
      // doc.setData({
      //   first_name: 'John',
      //   last_name: 'Doe',
      //   table:[
      //     {
      //       "one": "John",
      //       "two": "Doe",
      //       "three": "+44546546454",
      //       "four": "fsfds",
      //       "five": "fdfsdfs"
      //     },
      //     {
      //       "one": "Jane",
      //       "two": "Doe",
      //       "Three": "+445476454",
      //       "four": "fsfds",
      //       "five": "fdfsdfs"
      //     }
      //   ]
      // });
      try {
        // Replace placeholders with info values
        doc.render()
      } catch (error) {
        const e = {
          message: error.message,
          name: error.name,
          stack: error.stack,
          properties: error.properties,
        }
        console.log(JSON.stringify({error: e}))
        throw error
      }

      const out = doc.getZip().generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      })
      // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
      saveAs(out, 'output.docx')
    }

    reader.onerror = function() {
      console.log(reader.error)
    };
  }
}


