import { Component, ViewChild } from '@angular/core';

import { Http, RequestOptions, Headers, Response, ResponseContentType } from '@angular/http';
import { Observable } from 'rxjs/Rx';  
import * as XLSX from 'xlsx';
import { DataTable, DataTableResource } from './data-table';

import * as FileSaver from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'app';
  constructor(private http: Http) {
  }
  @ViewChild('inputFile')
  inputFile: any;
  data: any;
  filteredData: any;
  dataheader: any;
  orginalDataheader: any;
  dataHeaderJSON : any;
  selectedColumnIndex:any[] = [];
  count: number = -1;
  existedColumns : any[] = [{'label':'First Column','selected':false},{'label':'Second Column','selected':false},{'label':'Third Column','selected':false},{'label':'Fourth Column','selected':false}]
  wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
  fileName: string = 'SheetJS.xlsx';
  test: any[] = [];

  onFileChange(evt: any) {
		/* wire up file reader */
		this.count = -1;
		const target: DataTransfer = <DataTransfer>(evt.target);
		if (target.files.length !== 1) throw new Error('Cannot use multiple files');
		const reader: FileReader = new FileReader();
		reader.onload = (e: any) => {
			/* read workbook */
			const bstr: string = e.target.result;
			const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});

			/* grab first sheet */
			const wsname: string = wb.SheetNames[0];
			const ws: XLSX.WorkSheet = wb.Sheets[wsname];

			/* save data */
			this.dataheader = <any>(XLSX.utils.sheet_to_json(ws, {header: 1}));
			this.orginalDataheader = <any>(XLSX.utils.sheet_to_json(ws, {header: 1}));
			this.data = XLSX.utils.sheet_to_json(ws, {raw: true});
			this.dataHeaderJSON = [];
			(<any[]>this.dataheader[0]).forEach( item =>{
                this.dataHeaderJSON.push({'datafield':item,'title':item,'selected':'false'})
            })
		};
    reader.readAsBinaryString(target.files[0]);

    //Newly Added Code
   
    let file: File = target.files[0];
      let formData: FormData = new FormData();
      formData.append('uploadFile', file, file.name);
      let headers = new Headers()
      //headers.append('Content-Type', 'json');  
      //headers.append('Accept', 'application/json');  
      let options = new RequestOptions({ headers: headers });
    let apiUrl1 = "http://localhost:52797/api/UploadFileApi/UploadJsonFile";
      this.http.post(apiUrl1, formData, options)
        .map(res => res.json())
        .catch(error => Observable.throw(error))
        .subscribe(
          data => console.log('success'),
          error => console.log(error)
        )  

    //Ended Here

	}
  new(): void {
    console.log(this.inputFile.nativeElement.files);
    this.inputFile.nativeElement.value = "";
    console.log(this.inputFile.nativeElement.files);
    this.data = [];
    this.filteredData = [];
    this.dataheader = [];
    this.dataHeaderJSON = [];
    this.selectedColumnIndex = [];
    this.count = -1;    
  }

	//export() {
	//	///* generate worksheet */
	//	//const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.dataheader);

	//	///* generate workbook and add the worksheet */
	//	//const wb: XLSX.WorkBook = XLSX.utils.book_new();
	//	//XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

	//	///* save to file */
 // //    XLSX.writeFile(wb, this.fileName);



 //     let headers = new Headers({ 'Content-Type': 'application/json' });

 //     let options = new RequestOptions({ headers: headers });

 //     this.http.get("http://localhost:52797/api/UploadFileApi/Get", options)
 //       .map(res => res.json())
 //       .catch(error => Observable.throw(error))
 //       .subscribe(
 //         data => console.log('success'),
 //         error => console.log(error)
 //       )
     
  //}

  export(): void {
    let url: string = "http://localhost:52797/api/UploadFileApi/Get";
    let headers = new Headers({ 'Content-Type': 'application/json' });

    let options = new RequestOptions({ responseType: ResponseContentType.Blob, headers });

    this.http.get(url, options)
      .map(res => res.blob())
      .subscribe(
        data => {
          FileSaver.saveAs(data,"Export.xls");
        },
        err => {
          console.log('error');
          console.error(err);
        });
  }

	save(): void{
		this.filteredData = [];
		(<any[]>this.dataheader).forEach( item =>{
			let individualItem:any[] = [];
			for(let ob of this.selectedColumnIndex)
				individualItem.push(item[ob]);
			this.filteredData.push(individualItem);
		});
        this.dataheader = this.filteredData;

        let headers = new Headers({ 'Content-Type': 'application/json' });
       
        let options = new RequestOptions({ headers: headers });
     
      let values = JSON.stringify(this.test);
      this.test = [];
      this.http.post("http://localhost:52797/api/UploadFileApi/CopyExcel", values, options)
            .map(res => res.json())
            .catch(error => Observable.throw(error))
            .subscribe(
                data => console.log('success'),
                error => console.log(error)
            )  
	}

	reset():void {
		this.dataHeaderJSON = [];
		this.dataheader = this.orginalDataheader;
		console.log(this.dataheader[0]);
		this.dataheader[0].forEach( item =>{
					this.dataHeaderJSON.push({'datafield':item,'title':item,'selected':'false'})
	            })
		  this.selectedColumnIndex  = [];
		   this.filteredData = [];
		  this.count = -1;
		}
	itemHandler(selectedItem) {
		this.updateStyles(selectedItem);
		this.updateColumnName(selectedItem);
    }

    updateStyles(selectedItem){
        this.count++;
        if(this.count < this.existedColumns.length){
	        let temp = this.dataHeaderJSON;
			this.dataHeaderJSON = [];
			this.existedColumns[this.count].selected = true;
			(<any[]>temp).forEach( item =>{
				if(selectedItem.title==item.title){
		            this.dataHeaderJSON.push({'datafield':item.title,'title':this.existedColumns[this.count].label,'selected':'columnSelected'})
				}
				else{
					this.dataHeaderJSON.push({'datafield':item.datafield,'title':item.title,'selected':item.selected})
				}
	        })
	    }		
    }

    updateColumnName(selectedItem){
		for(var i=0;i<this.dataheader[0].length;i++){
          if (this.dataheader[0][i] == selectedItem.title) {
            this.test.push(this.dataheader[0][i]+":"+this.existedColumns[this.count].label);
				this.dataheader[0][i]=this.existedColumns[this.count].label;
				this.selectedColumnIndex.push(i);
			}
		}
    }
}
