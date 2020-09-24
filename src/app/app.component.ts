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

  loadFile(url, callback) {
    JSZipUtils.getBinaryContent(url, callback);
  }

  test(val) {
    let form = document.forms;
    console.log(form);
  }

  exportdocx(): void {
    this.loadFile('https://docxtemplater.com/tag-example.docx',(error, content) => {
      if (error) {
        throw error;
      }
      let zip = new JSZip(content);
      const doc = new Docxtemplater().loadZip(zip);
      doc.setData({
        first_name: 'John',
        last_name: 'Doe',
        phone: '0652455478',
        description: 'New Website'
      });
      try {
        // Replace placeholders with info values
        doc.render();
      } catch (error) {
        const e = {
          message: error.message,
          name: error.name,
          stack: error.stack,
          properties: error.properties,
        };
        console.log(JSON.stringify({error: e}));
        throw error;
      }

      const out = doc.getZip().generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });
        // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
      saveAs(out, 'output.docx');
    });
  }

  // public company = [
  //   { "ContainerNo": 4 },
  //   { "ContainerNo": 4, "SelCondition": 23, "ContainerCondition": 45,  "GateInDateTime": 55 },
  //   { "ContainerNo": 4, "SelCondition": 23, "ContainerCondition": 45,  "GateInDateTime": 55 }
  // ];
}


