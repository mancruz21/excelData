import { inject, Component } from '@angular/core';
import { Firestore, collection, getDocs } from '@angular/fire/firestore';
import * as XLSX from 'xlsx';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'downloadExcel';
  depto: string = '';

  collection: any;
  firestore: Firestore = inject(Firestore);

  constructor() {
    this.collection = collection(this.firestore, 'personasFirestore');
  }

  downloadExcel() {
    if(this.depto == ''){
      alert('Debe seleccionar un departamento.');
      return;
    }
    getDocs(this.collection)
      .then((querySnapshot) => {
        let data: any[] = [];
        querySnapshot.forEach((doc) => {
          data.push(doc.data());
        });
        
        // Crear un arreglo de filas para los datos
        let rowsPreview: any[] = [];
        
        // Crear la primera fila con encabezados
        const headers = [
          'tipoId',
          'id_document',
          'departamento',
          'Apellidos',
          'Nombres',
          'Sexo',
          'Embarazada',
          'FechaNacimiento',
          'Edad',
          'LugarNacimiento',
          'email',
          'lugarResidencia',
          'Vereda',
          'nombreFinca',
          'Predio',
          'RtaPredio',
          'extension',
          'Latitud',
          'Longitud',
          'Altura',
          'escolaridad',
          'Telefono',
          'estadoCivil',

          'regimenSalud',
          'conformanHogar',
          'numHombres',
          'numMujeres',
          'productor',
          'parentezco',
          'discapacidad',
          'rtaDiscapacidad',
          'etnia',
          'otraEtnia',
          'cabezaFlia',
          'internet',
          'SiInternet',
          'dispositivo',
          'tipoTelefono',
          'capacitaciones',
          'plataforma',
          'redes',
          
          'principalIngreso',
          'OtroIngreso',
          'produccionAños',
          'areaProduccion',
          'costoProduccion',
          'precioVenta',
          'kilogramos',
          'precioKilos',
          'vendio',
          'vendioCosecha',
          'Financia',
          'promedio',
          'rtaOtro',

          'edadPlantas',
          'AlturaPlantas',
          'TiempoProduc',
          'PlantasHectarea',
          'PersonasTrabajan',
          'Variedad',
          'Propagacion',
          'Semillas',
          'Cosecha_Principal_Inicial',
          'Cosecha_Principal_Final ',
          'Cosecha_Mitaca_Inicial ',
          'Cosecha_Mitaca_Final ',
          'AguaPotable ',
          'Energia ',
          'Agua ',
          'Analisis ',
          'Plagas ',
          'ManejaPlagas ',
          'quimicoCual  ',
          'frecuencia ',
          'cantidad  ',
          'abono  ',
          'abonocual  ',
          'abonofrecuencia  ',
          'abonocantidad  ',
          'PlagasDificil  ',
          'Cosecha  ',
          'proceso  ',
          'Herramientas  ',
          'HerramientasDesinfectadas  ',
          'OtroDesinfectante  ',
          'Desafio   ',
          'Transporte  ',
          'otroTransporte ',

          'Encuestador',
          'FechaDiligenciamiento',
          'departamento',
          'id_document',
          'tipoId',
        ];
        rowsPreview.push(headers);
        // Agregar los datos al arreglo de filas
        data.forEach((item) => {
          const row = [
            item['tipoId']?item['tipoId']:null,
            item['component1']?item['component1']['id_document']:null,
            item['component1']?item['component1']['departamento']:null,
            item['component1']?item['component1']['Apellidos']:null,
            item['component1']?item['component1']['Nombres']:null,
            item['component1']?item['component1']['Sexo']:null,
            item['component1']?item['component1']['Embarazada']:null,
            item['component1']?item['component1']['FechaNacimiento']:null,
            item['component1']?item['component1']['Edad']:null,
            item['component1']?item['component1']['LugarNacimiento']:null,
            item['component1']?item['component1']['email']:null,
            item['component1']?item['component1']['LugarResidencia']:null,
            item['component1']?item['component1']['Vereda']:null,
            item['component1']?item['component1']['nombreFinca']:null,
            item['component1']?item['component1']['Predio']:null,
            item['component1']?item['component1']['RtaPredio']:null,
            item['component1']?item['component1']['extension']:null,
            item['component1']?item['component1']['coordenadas']:null,
            item['component1']?item['component1']['latitud']:null,
            item['component1']?item['component1']['longitud']:null,
            item['component1']?item['component1']['Altura']:null,
            item['component1']?item['component1']['escolaridad']:null,
            item['component1']?item['component1']['Telefono']:null,
            item['component1']?item['component1']['estadoCivil']:null,

            item['component2']?item['component2']['regimenSalud']:null,
            item['component2']?item['component2']['conformanHogar']:null,
            item['component2']?item['component2']['numHombres']:null,
            item['component2']?item['component2']['numMujeres']:null,
            item['component2']?item['component2']['productor']:null,
            item['component2']?item['component2']['parentezco']:null,
            item['component2']?item['component2']['discapacidad']:null,
            item['component2']?item['component2']['rtaDiscapacidad']:null,
            item['component2']?item['component2']['etnia']:null,
            item['component2']?item['component2']['otraEtnia']:null,
            item['component2']?item['component2']['cabezaFlia']:null,
            item['component2']?item['component2']['internet']:null,
            item['component2']?item['component2']['SiInternet']:null,
            item['component2']?item['component2']['dispositivo']:null,
            item['component2']?item['component2']['tipoTelefono']:null,
            item['component2']?item['component2']['capacitaciones']:null,
            item['component2']?item['component2']['plataforma']:null,
            item['component2']?item['component2']['redes']:null,
           
            item['component3']?item['component3']['principalIngreso']:null,
            item['component3']?item['component3']['OtroIngreso']:null,
            item['component3']?item['component3']['produccionAños']:null,
            item['component3']?item['component3']['areaProduccion']:null,
            item['component3']?item['component3']['costoProduccion']:null,
            item['component3']?item['component3']['precioVenta']:null,
            item['component3']?item['component3']['kilogramos']:null,
            item['component3']?item['component3']['precioKilos']:null,
            item['component3']?item['component3']['vendio']:null,
            item['component3']?item['component3']['vendioCosecha']:null,
            item['component3']?item['component3']['Financia']:null,
            item['component3']?item['component3']['promedio']:null,
            item['component3']?item['component3']['rtaotro']:null,

            item['component4']?item['component4']['edadPlantas']:null,
            item['component4']?item['component4']['AlturaPlantas']:null,
            item['component4']?item['component4']['TiempoProduc']:null,
            item['component4']?item['component4']['PlantasHectarea']:null,
            item['component4']?item['component4']['PersonasTrabajan']:null,
            item['component4']?item['component4']['Variedad']:null,
            item['component4']?item['component4']['Propagacion']:null,
            item['component4']?item['component4']['Semillas']:null,
            item['component4']?item['component4']['Cosecha_Principal_Inicial']:null,
            item['component4']?item['component4']['Cosecha_Principal_Final']:null,
            item['component4']?item['component4']['Cosecha_Mitaca_Inicial']:null,
            item['component4']?item['component4']['Cosecha_Mitaca_Final']:null,
            item['component4']?item['component4']['AguaPotable']:null,
            item['component4']?item['component4']['Energia']:null,
            item['component4']?item['component4']['Agua']:null,
            item['component4']?item['component4']['Analisis']:null,
            item['component4']?item['component4']['Plagas']:null,
            item['component4']?item['component4']['ManejaPlagas']:null,
            item['component4']?item['component4']['quimicoCual']:null,
            item['component4']?item['component4']['cantidad']:null,
            item['component4']?item['component4']['abono']:null,
            item['component4']?item['component4']['abonocual']:null,
            item['component4']?item['component4']['abonofrecuencia']:null,
            item['component4']?item['component4']['abonocantidad']:null,
            item['component4']?item['component4']['PlagasDificil']:null,
            item['component4']?item['component4']['Cosecha']:null,
            item['component4']?item['component4']['proceso']:null,
            item['component4']?item['component4']['Herramientas']:null,
            item['component4']?item['component4']['HerramientasDesinfectadas']:null,
            item['component4']?item['component4']['OtroDesinfectante']:null,
            item['component4']?item['component4']['Desafio']:null,
            item['component4']?item['component4']['Transporte']:null,
            item['component4']?item['component4']['otroTransporte']:null,
            
            item['component5']?item['component5']['Encuestador']:null,
            item['component5']?item['component5']['FechaDiligenciamiento']:null,
            item['departamento']?item['departamento']:null,
            item['id_document']?item['id_document']:null,
            item['tipoId']?item['tipoID']:null,
          ];
          rowsPreview.push(row);
        });
        let rows = [];
        rows.push(rowsPreview[0]);
        rowsPreview.forEach((x) => {
          if(x[13]== this.depto){
            rows.push(x);
          }
        });
        
        // Crear un libro de Excel y una hoja de trabajo
        const wb: XLSX.WorkBook = XLSX.utils.book_new();
        const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(rows);
  
        // Agregar la hoja de trabajo al libro de Excel
        XLSX.utils.book_append_sheet(wb, ws, 'Hoja1');
  
        // Convertir el libro de Excel a un blob de datos
        const blobData = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
  
        // Crear un Blob a partir de la cadena base64
        const blob = new Blob([this.base64ToArrayBuffer(blobData)], {
          type:
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
  
        // Crear un objeto URL para el blob de datos
        const blobUrl = URL.createObjectURL(blob);
  
        // Crear un elemento <a> invisible para descargar el archivo
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = blobUrl;
        a.download = 'data.xlsx';
  
        // Agregar el elemento <a> al DOM y simular un clic para descargar el archivo
        document.body.appendChild(a);
        a.click();
  
        // Limpiar el objeto URL y eliminar el elemento <a> del DOM
        URL.revokeObjectURL(blobUrl);
        document.body.removeChild(a);
      })
      .catch((error) => {
        console.error('Error obteniendo documentos: ', error);
      });
  }
  
  
  // Función para convertir una cadena base64 a ArrayBuffer
  base64ToArrayBuffer(base64: string): ArrayBuffer {
    const binaryString = window.atob(base64);
    const length = binaryString.length;
    const bytes = new Uint8Array(length);
    for (let i = 0; i < length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
  }
  
}


