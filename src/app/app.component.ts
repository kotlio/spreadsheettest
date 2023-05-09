import { Component } from '@angular/core';
import * as XLSX from "xlsx" ;
import { WorkSheet } from 'xlsx';

declare var x_spreadsheet: any;

class ExportedCell{
    address: string = "";
    value: string = "";
}


@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'spreadSheetTest';
  fileToUpload: File | null = null;
  mySpreadSheet: any;
  exportedCells: ExportedCell[] = [];
  ngOnInit(){

  }

  public onExport_Click(){
    this.exportedCells = [];
    this.getExportedCells();
  }

  async onFileSelected(event: any){
    if (event != null){
      this.fileToUpload = event.target.files.item(0);

      const options = {
        mode: 'edit', // edit | read
        showToolbar: false,
        showGrid: true,
        showContextmenu: true,
        view: {
          height: () => 800,
          width: () => 1000,
        },
        row: {
          len: 100,
          height: 25,
        },
        col: {
          len: 26,
          width: 100,
          indexWidth: 60,
          minWidth: 60,
        },
        style: {
          bgcolor: '#ffffff',
          align: 'left',
          valign: 'middle',
          textwrap: false,
          strike: false,
          underline: false,
          color: '#0a0a0a',
          font: {
            name: 'Helvetica',
            size: 10,
            bold: false,
            italic: false,
          },
        },
      }


      this.mySpreadSheet = new x_spreadsheet('#xspreadsheet-demo', options);
      const arrayBuff = await this.fileToUpload?.arrayBuffer();
      var workbook = XLSX.read(arrayBuff);
      this.mySpreadSheet.loadData(this.stox(workbook));
    }

  } 

  stox(wb: XLSX.WorkBook) {
    var out: any = [];
    wb.SheetNames.forEach(function (name) {
      var o: WorkSheet = { name: name, rows: [], merges: [] };
      var ws = wb.Sheets[name];
      if(!ws || !ws["!ref"]) return;
      var range = XLSX.utils.decode_range(ws['!ref']);
      // sheet_to_json will lost empty row and col at begin as default
      range.s = { r: 0, c: 0 };
      var aoa = XLSX.utils.sheet_to_json(ws, {
        raw: false,
        header: 1,
        range: range
      });
  
      aoa.forEach(function (r: any, i: any) {
        var cells: any = {};
        r.forEach(function (c: any, j: any) {
          let buffEditable = j > 2 ? true : false;
          cells[j] = { text: c, editable: buffEditable };
  
          var cellRef = XLSX.utils.encode_cell({ r: i, c: j });
  
          if ( ws[cellRef] != null && ws[cellRef].f != null) {
            cells[j].text = "=" + ws[cellRef].f;
          }
        });
        o['rows'][i] = { cells: cells };
      });
  
      o['merges'] = [];
      (ws["!merges"]||[]).forEach(function (merge, i) {
        //Needed to support merged cells with empty content
        if (o['rows'][merge.s.r] == null) {
          o['rows'][merge.s.r] = { cells: {} };
        }
        if (o['rows'][merge.s.r].cells != undefined) {
          if (o['rows'][merge.s.r].cells[merge.s.c] == null){
            o['rows'][merge.s.r].cells[merge.s.c] = {};

            o['rows'][merge.s.r].cells[merge.s.c].merge = [
              merge.e.r - merge.s.r,
              merge.e.c - merge.s.c
            ];
          }
        }
  
        o['rows'][i] = XLSX.utils.encode_range(merge);
      });
  
      out.push(o);
    });
  
    return out;
  }

  getExportedCells(){
    let cell;
    let cellStyle;
    for (let ri = 0; ri < this.mySpreadSheet.options.row.len; ri++){
      for (let ci = 0; ci < this.mySpreadSheet.options.col.len; ci++){
        cell = this.mySpreadSheet.cell(ri, ci);
        cellStyle = this.mySpreadSheet.cellStyle(ri, ci);

        if (cell != null && cell['to-export'] === true){
          let newExportCell = new ExportedCell();
          newExportCell.address = "строка №" + (ri + 1).toString() + ',' + "столбец №" + (ci+1).toString();
          newExportCell.value = cell.text;
          this.exportedCells.push(newExportCell);
        }

      }
    }
  }
}
