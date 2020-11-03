import { Component, ElementRef, ViewChild } from '@angular/core';  
import * as XLSX from 'xlsx'; 
 
@Component({  
  selector: 'app-root',  
  templateUrl: './app.component.html',  
  styleUrls: ['./app.component.scss']  
})  
export class AppComponent {  
  @ViewChild('TABLE', { static: false }) TABLE: ElementRef;  
  title = 'Excel';  
  ExportTOExcel() {  
    const ws: XLSX.WorkSheet = XLSX.utils.table_to_sheet(this.TABLE.nativeElement);  
    const wb: XLSX.WorkBook = XLSX.utils.book_new();  
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');  
    XLSX.writeFile(wb, 'ScoreSheet.xlsx');  
  }  

  header = [
        {
            "displayName": "Vision OUC",
            "key": "Vision_ouc",
            "pivotOption": "R"
        },
        {
            "displayName": "Period",
            "key": "Period",
            "pivotOption": "C"
        }, {
            "displayName": "Avg Balance",
            "key": "balance",
            "pivotOption": "V "
        }
    ]
    body = [
          {
              "Vision_ouc": "KE01000017",
              "Period": "20160401",
              "balance": "163164849.07171"
          },
          {
              "Vision_ouc": "KE01000017",
              "Period": "20160501",
              "balance": "162366406.0121"
          },
          {
              "Vision_ouc": "KE01000018",
              "Period": "20160401",
              "balance": "347325574.38977"
          },
          {
              "Vision_ouc": "KE01000018",
              "Period": "20160501",
              "balance": "346564126.31549"
          },
          {
              "Vision_ouc": "KE01000019",
              "Period": "20160401",
              "balance": "183543736.37852"
          },
          {
              "Vision_ouc": "KE01000019",
              "Period": "20160501",
              "balance": "186111804.54701"
          }
      ]


    }  