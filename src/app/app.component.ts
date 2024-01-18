import { Component } from '@angular/core';
import { FormControl } from '@angular/forms';
import { FormGroup } from '@angular/forms';
////
import { ApirestService } from './providers/apirest.service';
import { ExcelService } from './providers/excel.service';
///

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
  providers:[ApirestService]
})
export class AppComponent {
  title = 'dialog2';

  visible : boolean = true;
  nombre : string = "";
  form = new FormGroup({
    nombre : new FormControl('')
  });

///////////////////////// EXCEL /////////////////////////////////
  constructor( private api : ApirestService, private excel : ExcelService){}
  descargar(){
    this.api.getProducts().subscribe((resp)=> {
      console.log(resp)
      this.excel.descargarExcel(resp); //pasando datos de la base de datos
    })
  }

}
