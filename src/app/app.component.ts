import { Component } from '@angular/core';
import * as XLSX  from 'xlsx';
import { JsonSheet } from './model/model';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
@Component({
	selector: 'app-root',
	templateUrl: './app.component.html',
	styleUrls: ['./app.component.css']
})
export class AppComponent {
	title = 'app';
	file:File
	data:JsonSheet[] = null;
	content: string;
	onFileSelected(event) {
		this.file = event.target.files[0];
	}
	
	upload() {
		const reader: FileReader = new FileReader();
		reader.readAsBinaryString(this.file);
		reader.onload = (e: any) => {
			console.log('e', e)
		  /* create workbook */
		  const binarystr: string = e.target.result;
		  const wb: XLSX.WorkBook = XLSX.read(binarystr, { type: 'binary' });
	
		  /* selected the first sheet */
		  const wsname: string = wb.SheetNames[0];
		  const ws: XLSX.WorkSheet = wb.Sheets[wsname];
	
		  /* save data */
		  
		  this.data = XLSX.utils.sheet_to_json<JsonSheet>(ws); // to get 2d array pass 2nd parameter as object {header: 1}
		  console.log(this.data);
		  this.genarateTable();
		}
	}

	export() {
		const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(this.data);
		const worksheet2: XLSX.WorkSheet = XLSX.utils.json_to_sheet(this.sheet2(this.data));
		const worksheet3: XLSX.WorkSheet = XLSX.utils.json_to_sheet(this.sheet3(this.data));
		const workbook: XLSX.WorkBook = {
		  Sheets: { Sheet1: worksheet,Sheet2:worksheet2,Sheet3:worksheet3 },
		  SheetNames: ['Sheet1','Sheet2','Sheet3'],
		};
	
		const excelBuffer: any = XLSX.write(workbook, {
		  bookType: 'xlsx',
		  type: 'array',
		});
	
		const data: Blob = new Blob([excelBuffer], { type: EXCEL_TYPE });
		const date = new Date();
		const fileName = `example${date.getTime()}.xlsx`;
	
		// FileSaver.saveAs(data, fileName);
		const link = document.createElement("a");
		link.href = URL.createObjectURL(data);
		link.download = fileName;
		link.click();
		link.remove();
	}

	genarateTable(){
		const display = document.getElementsByClassName("display");
		let header:string[]=  Object.keys(this.data[0]);
		// console.log('header', header)
		let th = "";
		header.map(i => {
			th += `<th>${i}</th>`
		});
		const tr = `<tr>${th}</tr>`
		let body = "";
		this.data.map(i => {
			body += `
			<tr>
				<td>${i.index}</td>
				<td>${i.value}</td>
			</tr>`
		})
		const table = `<table>${tr} ${body}</table>`;
		this.content = table;
		
	}

	sheet2(data:JsonSheet[]){
		return data.filter((val,index) => (index % 2) == 0);

	}

	sheet3(data:JsonSheet[]){
		return data.filter((val,index) => (index % 2) != 0);
	}
}
