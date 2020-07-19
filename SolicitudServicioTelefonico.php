<!DOCTYPE html>
<html lang="en" dir="ltr">
  <head>
    <meta charset="utf-8">
    <title>SolicitudServicioTelefonico</title>
    <link rel="stylesheet"  href="/formulario/css/prueba.css">
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
  </head>
  <body>
    <?php
    require 'class/vendor/autoload.php';
    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
    $usersheet= \PhpOffice\PhpSpreadsheet\IOFactory::load('UserFO.xlsx');
      $userworksheet=$usersheet->setActiveSheetIndexByName('UserFO');
      $usermax = $userworksheet->getHighestRow();
      $workflow = array();
      $workflow[0]="Solicitante";
      $workflow[1]="Gerente";
      $workflow[2]="Gerente General";
      $workflow[3]="Analista IT";
      $workflow[4]="Jefe IT";
      $workflow[5]="Gerente IT";
      $userarray=Array();
      $id_rolearray=Array();
      $aux=0;
      for ($i=2; $i <=$usermax ; $i++) {
        $id_usersaux=$userworksheet->getCell("A".$i);
        $id_user=$id_usersaux->getValue();
        $id_roleaux=$userworksheet->getCell("E".$i);
        $id_role=$id_roleaux->getValue();
        $id_rolearray[$aux]=$id_role;
        $userarray[$aux]=$id_user;
        $aux++;
      }
    $spreadsheet= \PhpOffice\PhpSpreadsheet\IOFactory::load('SolicitudHardwareSoftwareFO.xlsx');
    $worksheet=$spreadsheet->setActiveSheetIndexByName('SolicitudServicioTelefonico');
    $max = $worksheet->getHighestRow();
    $id_solicitudaux=$worksheet->getCell("C".$max);
    $id_solicitud=$id_solicitudaux->getValue();
    if ($id_solicitud=='id_solicitud') {
      $id_solicitud=1;
    }
    else {
      $id_solicitud++;
    }
    if($_POST){
    for ($i=0; $i <6 ; $i++) {
      $worksheet->getCell('B'.($max+$i+1))->setValue($i+1);
      $worksheet->getCell('C'.($max+$i+1))->setValue($_POST['id_solicitud']);
      $worksheet->getCell('D'.($max+$i+1))->setValue($_POST['Cambiar']);
      $worksheet->getCell('E'.($max+$i+1))->setValue($_POST['Asunto']);
      $worksheet->getCell('F'.($max+$i+1))->setValue($_POST['Centro']);
      $worksheet->getCell('G'.($max+$i+1))->setValue($_POST['Ubicacion']);
      $worksheet->getCell('H'.($max+$i+1))->setValue($_POST['Email']);
      $worksheet->getCell('I'.($max+$i+1))->setValue($_POST['TipoSoli']);
      $worksheet->getCell('J'.($max+$i+1))->setValue($_POST['TipoServicio']);
      $worksheet->getCell('K'.($max+$i+1))->setValue($_POST['Justificacion']);
      $worksheet->getCell('N'.($max+$i+1))->setValue($_POST['Marca']);
      $worksheet->getCell('O'.($max+$i+1))->setValue($_POST['Modelo']);
      $worksheet->getCell('P'.($max+$i+1))->setValue($_POST['NumeroDeSerie']);
      $worksheet->getCell('Q'.($max+$i+1))->setValue($_POST['Fecha']);

    }
    $worksheet->getCell('A'.($max+1))->setValue($_POST['id_user1']);
      $worksheet->getCell('A'.($max+2))->setValue($_POST['id_user2']);
      $worksheet->getCell('A'.($max+3))->setValue($_POST['id_user3']);
      $worksheet->getCell('A'.($max+4))->setValue($_POST['id_user4']);
      $worksheet->getCell('A'.($max+5))->setValue($_POST['id_user5']);
      $worksheet->getCell('A'.($max+6))->setValue($_POST['id_user6']);

    $worksheet->getCell('L'.($max+1))->setValue($_POST['Comentario1']);
    $worksheet->getCell('M'.($max+1))->setValue($_POST['estado1']);
    $worksheet->getCell('L'.($max+2))->setValue($_POST['Comentario2']);
    $worksheet->getCell('M'.($max+2))->setValue($_POST['estado2']);
    $worksheet->getCell('L'.($max+3))->setValue($_POST['Comentario3']);
    $worksheet->getCell('M'.($max+3))->setValue($_POST['estado3']);
    $worksheet->getCell('L'.($max+4))->setValue($_POST['Comentario4']);
    $worksheet->getCell('M'.($max+4))->setValue($_POST['estado4']);
    $worksheet->getCell('L'.($max+5))->setValue($_POST['Comentario5']);
    $worksheet->getCell('M'.($max+5))->setValue($_POST['estado5']);
    $worksheet->getCell('L'.($max+6))->setValue($_POST['Comentario6']);
    $worksheet->getCell('M'.($max+6))->setValue($_POST['estado6']);


    $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save('SolicitudHardwareSoftwareFO.xlsx');
    header("Location: /formulario/index.html");
    }
    ?>
    <a title="toyota" href="/formulario/index.html"><img src="/formulario/toyota.png" alt=""></a>
    <br>
    <form class="" action="/formulario/SolicitudServicioTelefonico.php" method="post">
      <label for="Tipo">Tipo de solicitud</label>
      <select name="Tipo">
       <option>SolicitudServicioTelefonico</option>
      </select>
      <br>
      <label for="id_user1">Id_user para Solicitante</label>
          <select class="" name="id_user1">
            <?php for ($indice=0; $indice <=count($userarray) ; $indice++) {?>
              <?php  if($id_rolearray[$indice]==1){?>
                <option value="<?php echo  $userarray[$indice]; ?>"><?php echo $userarray[$indice]; ?></option>
              <?php }}?>
          </select>
          <label for="id_user2">Id_user para Gerente</label>
          <select class="" name="id_user2">
            <?php for ($indice=0; $indice <=count($userarray) ; $indice++) {?>
              <?php  if($id_rolearray[$indice]==2){?>
                <option value="<?php echo  $userarray[$indice]; ?>"><?php echo $userarray[$indice]; ?></option>
              <?php }}?>
          </select>
          <label for="id_user3">Id_user para Gerente General</label>
          <select class="" name="id_user3">
            <?php for ($indice=0; $indice <=count($userarray) ; $indice++) {?>
              <?php  if($id_rolearray[$indice]==3){?>
                <option value="<?php echo  $userarray[$indice]; ?>"><?php echo $userarray[$indice]; ?></option>
              <?php }}?>
          </select>
          <label for="id_user4">Id_user para Analista IT</label>
          <select class="" name="id_user4">
            <?php for ($indice=0; $indice <=count($userarray) ; $indice++) {?>
              <?php  if($id_rolearray[$indice]==4){?>
                <option value="<?php echo  $userarray[$indice]; ?>"><?php echo $userarray[$indice]; ?></option>
              <?php }}?>
          </select>
          <label for="id_user5">Id_user para Jefe IT</label>
          <select class="" name="id_user5">
            <?php for ($indice=0; $indice <=count($userarray) ; $indice++) {?>
              <?php  if($id_rolearray[$indice]==5){?>
                <option value="<?php echo  $userarray[$indice]; ?>"><?php echo $userarray[$indice]; ?></option>
              <?php }}?>
          </select>
          <label for="id_user6">Id_user para Gerente IT</label>
          <select class="" name="id_user6">
            <?php for ($indice=0; $indice <=count($userarray) ; $indice++) {?>
              <?php  if($id_rolearray[$indice]==6){?>
                <option value="<?php echo  $userarray[$indice]; ?>"><?php echo $userarray[$indice]; ?></option>
              <?php }}?>
          </select>
      <br>
      <label for="id_solicitud">id_solicitud</label>
        <select class="" name="id_solicitud">
              <option value="<?php echo $id_solicitud; ?>"><?php echo $id_solicitud; ?></option>
      </select>
      <br>
      <label for="Cambiar">Cambiar solicitante</label>
      <input type="text" id="Cambiar" name="Cambiar" value="">
      <br>
      <label for="Asunto">Asunto</label>
      <input type="text" id="Asunto" name="Asunto" value="">
      <br>
      <label for="Centro">Centro de costo(*)</label>
      <input type="text" id="Centro" name="Centro" value="">
      <br>
      <label for="Ubicacion">Ubicacion Geografica(*)</label>
      <input type="text" id="Ubicacion" name="Ubicacion" value="">
      <br>
      <label for="Email">Email(*)</label>
      <input type="text" id="Email" name="Email" value="">
      <br>
      <label for="TipoSoli">Tipo de solicitud</label>
      <input type="text" id="TipoSoli" name="TipoSoli" value="">
      <br>
      <label for="TipoServicio">Tipo de Servicio</label>
      <input type="text" id="TipoServicio" name="TipoServicio" value="">
      <br>
      <label for="Justificacion">Justificacion(*)</label>
      <input type="text" id="Justificacion" name="Justificacion" value="">
      <br>
      <label for="Comentario1">Comentario Estado Solicitud para Solicitante</label>
      <input type="text" id="Comentario1" name="Comentario1" value="">
      <label for="estado1">Estado solicitud para Solicitante</label>
      <select class="" name="estado1">
        <option value="Aprobada">Aprobada</option>
        <option value="Necesita correccion">Necesita correccion</option>
        <option value="Rechazado">Rechazado</option>
      </select>
      <br>
      <label for="Comentario2">Comentario Estado Solicitud para Gerente</label>
      <input type="text" id="Comentario2" name="Comentario2" value="">
      <label for="estado2">Estado para Gerente</label>
      <select class="" name="estado2">
        <option value="Aprobada">Aprobada</option>
        <option value="Necesita correccion">Necesita correccion</option>
        <option value="Rechazado">Rechazado</option>
      </select>
      <br>
      <label for="Comentario3">Comentario Estado Solicitud para Gerente General</label>
      <input type="text" id="Comentario3" name="Comentario3" value="">
      <label for="estado3">Estado para Gerente General</label>
      <select class="" name="estado3">
        <option value="Aprobada">Aprobada</option>
        <option value="Necesita correccion">Necesita correccion</option>
        <option value="Rechazado">Rechazado</option>
      </select>
      <br>
      <label for="Comentario4">Comentario Estado Solicitud para Analista IT</label>
      <input type="text" id="Comentario4" name="Comentario4" value="">
      <label for="estado4">Estado para Analista IT</label>
      <select class="" name="estado4">
        <option value="Aprobada">Aprobada</option>
        <option value="Necesita correccion">Necesita correccion</option>
        <option value="Rechazado">Rechazado</option>
      </select>
      <br>
      <label for="Comentario5">Comentario Estado Solicitud para Jefe IT</label>
      <input type="text" id="Comentario5" name="Comentario5" value="">
      <label for="estado5">Estado para Jefe IT</label>
      <select class="" name="estado5">
        <option value="Aprobada">Aprobada</option>
        <option value="Necesita correccion">Necesita correccion</option>
        <option value="Rechazado">Rechazado</option>
      </select>
      <br>
      <label for="Comentario6">Comentario Estado Solicitud para Gerente IT</label>
      <input type="text" id="Comentario6" name="Comentario6" value="">
      <label for="estado6">Estado para Gerente IT</label>
      <select class="" name="estado6">
        <option value="Aprobada">Aprobada</option>
        <option value="Necesita correccion">Necesita correccion</option>
        <option value="Rechazado">Rechazado</option>
      </select>
      <label for="Marca">Especificacion tecnica Marca</label>
      <input type="text" id="Marca" name="Marca" value="">
      <br>
      <label for="Modelo">Especificacion tecnica Modelo</label>
      <input type="text" id="Modelo" name="Modelo" value="">
      <br>
      <label for="NumeroDeSerie">Especificacion tecnica Numero de Serie</label>
      <input type="text" id="NumeroDeSerie" name="NumeroDeSerie" value="">
      <br>
      <label for="Fecha">Especificacion tecnica Fecha</label>
      <input type="text" id="Fecha" name="Fecha" value="">
      <br>
      <br>

      <button type="submit" class="btn btn-primary mb-2">Enviar</button>
    </form>

  </body>
</html>
