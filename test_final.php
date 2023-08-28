<!DOCTYPE html>
<html lang="en">
    <head>
        <title>Modificar Word</title>
      </head>
      <body>
        <h1>Modificar Word</h1>
        <form action="test_final.php" method="GET"><!-- Formulario para pasar las rutas-->
          <label for="ruta_archivo">Ruta del archivo:</label>
          <input type="text" name="ruta_archivo" id="ruta_archivo"><br><br>
          <label for="json_contenido">Contenido JSON:</label>
          <input type="text" name="json_contenido" id="json_contenido"><br><br>
          <input type="submit" value="modificar_Word">
        </form>
      </body>
</html>
<?php
  require_once 'C:\xampp\htdocs\word\vendor\autoload.php';
  use PhpOffice\PhpWord\TemplateProcessor;
  if(isset($_GET['ruta_archivo']) && isset($_GET['json_contenido'])) { //mientras haya contenido tanto en el imput de la ruta del word y del json llama a la funcion
      $ruta_archivo = $_GET['ruta_archivo']; //guarda en esta variable la ruta que le pongas en el imput correspondiente
      $json = $_GET['json_contenido'];//guarda en esta variable la ruta que le pongas en el imput correspondiente
      modificar_word($ruta_archivo, $json);//llama a la funcion con las variables que guardan las rutas (no le he puesto ninguna validacion, asi que si le metes rutas malas, no funciona)
  }
  function modificar_word($ruta_archivo, $json) {//aqui la funcion
      echo "HAS ENTRADO EN LA FUNCION!!, mira en tu carpeta donde tengas test_final.php para ver el archivo 'test_esteban_modificado.docx' en el cual podras ver los resultados!!";
      echo "<br>";
      $documento = new TemplateProcessor($ruta_archivo);//cargo la plantilla de word especifica (en este caso lo paso por parametro)
      $json_ruta = file_get_contents($json); //json_ruta es donde leo el contenido de la ruta del json que le pase, da iugal que json sea mientras sea un archivo json "logico" para este caso 
      $json_php = json_decode($json_ruta, true);//traigo cualquier json y lo guardo en una variable (en este caso lo paso por parametro)
      $variables = array('imgfirma', 'company_name', 'comercial_name', 'vat', 'legal_address', 'legal_city', 'legal_country', 'legal_state', 'legal_zip_code', 'notification_address', 'notification_city', 'notification_state', 'notification_country', 'notification_zip_code', 'representative_name', 'identity_num', 'job_position', 'iban', 'swift', 'industry', 'manufacturer', 'nato_holder', 'nato_certificate_level', 'services_summary', 'key_customers', 'main_projects', 'main_contact_name', 'main_contact_position', 'main_contact_prefix', 'main_contact_phone', 'main_contact_email', 'sales_contact_name', 'sales_contact_position', 'sales_contact_prefix', 'sales_contact_phone', 'sales_contact_email', 'technical_contact_name', 'technical_contact_position', 'technical_contact_prefix', 'technical_contact_phone', 'technical_contact_email', 'administration_contact_name', 'administration_contact_position', 'administration_contact_prefix', 'administration_contact_phone', 'administration_contact_email', 'quality_contact_name', 'quality_contact_position', 'quality_contact_prefix', 'quality_contact_phone', 'quality_contact_email', 'qrcodefir');
      $resultados = array($json_php['imgfirma'], $json_php['company_name'], $json_php['comercial_name'], $json_php['vat'], $json_php['legal_address'], $json_php['legal_city'], $json_php['legal_country'], $json_php['legal_state'], $json_php['legal_zip_code'], $json_php['notification_address'], $json_php['notification_city'], $json_php['notification_state'], $json_php['notification_country'], $json_php['notification_zip_code'], $json_php['representative_name'], $json_php['identity_num'], $json_php['job_position'], $json_php['iban'], $json_php['swift'], $json_php['industry'], $json_php['manufacturer'], $json_php['nato_holder'], $json_php['nato_certificate_level'], $json_php['services_summary'], $json_php['key_customers'], $json_php['main_projects'], $json_php['main_contact_name'], $json_php['main_contact_position'], $json_php['main_contact_prefix'], $json_php['main_contact_phone'], $json_php['main_contact_email'], $json_php['sales_contact_name'], $json_php['sales_contact_position'], $json_php['sales_contact_prefix'], $json_php['sales_contact_phone'], $json_php['sales_contact_email'], $json_php['technical_contact_name'], $json_php['technical_contact_position'], $json_php['technical_contact_prefix'], $json_php['technical_contact_phone'], $json_php['technical_contact_email'], $json_php['administration_contact_name'], $json_php['administration_contact_position'], $json_php['administration_contact_prefix'], $json_php['administration_contact_phone'], $json_php['administration_contact_email'], $json_php['quality_contact_name'], $json_php['quality_contact_position'], $json_php['quality_contact_prefix'], $json_php['quality_contact_phone'], $json_php['quality_contact_email'], $json_php['qrcodefir']);
      //con $json_php['key'] obtengo el valor de cada key por orden que tenga cualquier json con este mismo formato
      $documento->setValue($variables, $resultados);//le hago el setValue a todas las variables que coincidan en el word
      $documento->saveAs('test_esteban_modificado.docx');//creo la nueva plantilla modificada de la plantilla base
  }
?>