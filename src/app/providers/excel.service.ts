import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs'; /// npm i exceljs
import * as fs from 'file-saver';   // npm i file-saver    npm i -D @types/file-saver
import { productos } from './api_interface';

@Injectable({
  providedIn: 'root'
})
export class ExcelService {

  private workbook!: Workbook;

  private crearHoja(productos: productos[]){ // productos del tipo productos[](interface)

    const hoja = this.workbook.addWorksheet('Informe'); // hoja de excel y su nombre, en este caso Informe 

    // anchos de columna
    hoja.getColumn("A").width=10;
    hoja.getColumn("B").width=30;
    hoja.getColumn("C").width=23;
    hoja.getColumn("D").width=20;
    hoja.getColumn("E").width=20;
    //etc

    // unir columnas
    hoja.mergeCells("A1","B1")

    // formato currency
    const numFmtStr = '$#,##0_)'
    hoja.getColumn("D").numFmt=numFmtStr

    // alinear textos
    hoja.columns.forEach((column)=>{
      column.alignment={vertical:'middle', horizontal: 'left', wrapText: true} //alinear en el centro
    });


    ////////////// IMAGEN ////////////////
    //const logoId=this.workbook.addImage({
    //  base64: imgBase64aqui,
    //  extension: 'jpeg',

    //});
    
    //hoja.addImage(logoId);
    ///////////////////////////////////////////////////
    
    // alto de fila 1 
    hoja.getRow(1).height=20;
    
    // Titulo 
    const celdaTitulo = hoja.getCell('A1'); 
    celdaTitulo.value="Informe de inspección";
    celdaTitulo.style.font={bold: true, size: 14, name: 'Calibri'}; // bold=negritas , size= tamaño letra
    
    // Suma de pagos (Total)
    hoja.getCell('E1').value="Total de pagos: "
    hoja.getCell('E1').alignment={horizontal: 'right'}
    hoja.getCell('E1').font={bold: true, size: 12, name: 'Calibri'};
    
    hoja.getCell('E1').border = {
      top: {style:'thick', color: {argb:'000000'}},
      left: {style:'thick', color: {argb:'000000'}},
      bottom: {style:'thick', color: {argb:'000000'}},
    };

    
    const totalPagos = productos.reduce((acc, item)=> acc + item.price,0)
    const celdaTotalPagos = hoja.getCell('F1');
    celdaTotalPagos.value=totalPagos;
    celdaTotalPagos.font={bold: true, size: 12, name: 'Calibri'};
    celdaTotalPagos.numFmt=numFmtStr
    celdaTotalPagos.alignment={horizontal: 'left'}
    
    celdaTotalPagos.border = {
      top: {style:'thick', color: {argb:'000000'}},
      bottom: {style:'thick', color: {argb:'000000'}},
      right: {style:'thick', color: {argb:'000000'}}
    };

    //titulos de cabeceras
    const filaCabeceras = hoja.getRow(3); //fila de las cabeceras
    
    filaCabeceras.values= [
    "Id", 
    "Titulo", 
    "Descripción",  //columnas de cabeceras
    "Precio"] ;

    filaCabeceras.font={bold: true, size: 12, name: 'Calibri'};
    
   
    
    // insertar datos en columnas ////////////////
    const filas = hoja.getRows(4, productos.length)!;
      filas.forEach((row)=>{
        row.height=40            // formato filas con datos
      })
    for (let index = 0; index < filas.length; index++) {
      const itemData = productos[index]
      const row = filas[index]
      row.values=[
        itemData.id,           //A4, A5 ...
        itemData.title,        //B4. B5
        itemData.description,  //C4, C5        //columnas de cabeceras
        itemData.price                          // Datos a insertar
      ]

      //Conteo de filas (numero de partes)
      const celdaTotalFilas = hoja.getCell('C1');
      celdaTotalFilas.value="Total partes: "+filas.length;
      celdaTotalFilas.font={bold: true, size: 12, name: 'Calibri'};
      celdaTotalFilas.border = {
        top: {style:'thick', color: {argb:'000000'}},
        left: {style:'thick', color: {argb:'000000'}},
        bottom: {style:'thick', color: {argb:'000000'}},  // bordes
        right: {style:'thick', color: {argb:'000000'}}
    };
      
      //// color de celdas
      celdaTotalFilas.fill={
        type: "pattern",
	      pattern: "solid",
	      fgColor: {
		    argb: "87ebff"
	      },
	      bgColor: {
		    argb: "000000"
	      }
      }

      
      
    }
    

  }

  descargarExcel(dataExcel:any){
    this.workbook = new Workbook();
    this.workbook.creator='Insico';     //metadata creador

    this.crearHoja(dataExcel.products); // ocupar metodo para crear la hoja

    this.workbook.xlsx.writeBuffer().then((data)=>{ 
      const blob = new Blob([data])
      fs.saveAs(blob, "Informe_insp.xlsx"); //nombre del archivo excel
    })
  }

}
