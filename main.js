var XlsxTemplate = require('xlsx-template');

const fs = require("fs");
const path = require("path");
async function  GenerarReporte(datos) {
    try{
    await fs.readFile(path.join(__dirname, 'formato.xlsx'), async function(err, data) {
  
      // Create a template
      var template = await new XlsxTemplate(data);
  
      // Replacements take place on first sheet
      var sheetNumber = 1;
  
      // Set up some placeholder values matching the placeholders in the template
      
      var values = {
             
              
                vivienda_nombre:"nombre de la vivienda",
                vivienda_total_registros:20132.25,
                vivienda_consumo:"120KW",
                sensor_1_25:1,
                sensor_2_25:1,
                sensor_3_25:1,
                sensor_4_25:1,
                sensor_1_50:1,
                sensor_2_50:1,
                sensor_3_50:1,
                sensor_4_50:1,
                sensor_1_75:1,
                sensor_2_75:1,
                sensor_3_75:1,
                sensor_4_75:1,
                sensor_1_min:1,
                sensor_2_min:1,
                sensor_3_min:1,
                sensor_4_min:1,
                sensor_1_mean:1,
                sensor_2_mean:1,
                sensor_3_mean:1,
                sensor_4_mean:1,
                sensor_1_max:1,
                sensor_2_max:1,
                sensor_3_max:1,
                sensor_4_max:1,
                sensor_1_std:1,
                sensor_2_std:1,
                sensor_3_std:1,
                sensor_4_std:1,
                sensor_1_count:1,
                sensor_2_count:1,
                sensor_3_count:1,
                sensor_4_count:1,
                sensor_1_total:1,
                sensor_2_total:1,
                sensor_3_total:1,
                sensor_4_total:1,
                vivienda_1_25:2,
                vivienda_1_50:2,
                vivienda_1_75:2,
                vivienda_1_min:2,
                vivienda_1_mean:2,
                vivienda_1_max:2,
                vivienda_1_std:2,
                vivienda_count:2,
                vivienda_total:2,
                recomendacion_1:" Recomendacion..",
                recomendacion_2:" Recomendacion..",
                recomendacion_3:" Recomendacion..",
                recomendacion_4:" Recomendacion..",
                recomendacion_5:" Recomendacion..",
                recomendacion_6:" Recomendacion..",
                recomendacion_7:" Recomendacion..",
                recomendacion_8:" Recomendacion..",
                recomendacion_9:" Recomendacion..",
                recomendacion_10:" Recomendacion..",
                recomendacion_11:" Recomendacion..",
                recomendacion_12:" Recomendacion..",
                alerta_1:"Alerta ..",
                alerta_2:"Alerta ..",
                alerta_3:"Alerta ..",
                alerta_4:"Alerta ..",
                alerta_5:"Alerta ..",
                alerta_6:"Alerta ..",
                alerta_7:"Alerta ..",
                alerta_8:"Alerta ..",
                alerta_9:"Alerta ..",
                alerta_10:"Alerta ..",
                alerta_11:"Alerta ..",
                alerta_12:"Alerta ..",
                fecha:"27/04/2023",
                


                

               
                
  
              
          };
        
      // Perform substitution
      await template.substitute(sheetNumber, values);
  
      // Get binary data
      var date =  await template.generate({ type: 'nodebuffer',compression: "DEFLATE"});
  
  
      await fs.writeFile(path.join(__dirname, 'resultado.xlsx'), date, function(err) {
          if(err) {
            return console.log(err);
          }
          
      })
  }
    )
  
    return
  
    }catch(error){
      console.log(error)
    }
  
  }

  GenerarReporte("hola")