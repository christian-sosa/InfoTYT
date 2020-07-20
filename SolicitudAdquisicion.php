  <!DOCTYPE html>
  <html lang="en" dir="ltr">
    <head>
      <meta charset="utf-8">
      <title>SolicitudAdquisicion</title>
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
      $rolsheet= \PhpOffice\PhpSpreadsheet\IOFactory::load('RolesItFO.xlsx');
      $rolworksheet=$rolsheet->setActiveSheetIndexByName('RolesIT');
      $rolmax= $rolworksheet->getHighestRow();
      $spreadsheet= \PhpOffice\PhpSpreadsheet\IOFactory::load('SolicitudHardwareSoftwareFO.xlsx');
      $worksheet=$spreadsheet->setActiveSheetIndexByName('SolicitudAdquisicion');
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
        $worksheet->getCell('F'.($max+$i+1))->setValue($_POST['Detalle']);
        $worksheet->getCell('G'.($max+$i+1))->setValue($_POST['Justificacion']);
        $worksheet->getCell('H'.($max+$i+1))->setValue($_POST['Informacion']);
      }
      $rolworksheet->getCell('A'.($rolmax+1))->setValue($_POST[id_user4]);
      $rolworksheet->getCell('A'.($rolmax+2))->setValue($_POST[id_user5]);
      $rolworksheet->getCell('A'.($rolmax+3))->setValue($_POST[id_user6]);
      $rolworksheet->getCell('B'.($rolmax+1))->setValue(4);
      $rolworksheet->getCell('B'.($rolmax+2))->setValue(5);
      $rolworksheet->getCell('B'.($rolmax+3))->setValue(6);
      $rolworksheet->getCell('C'.($rolmax+1))->setValue(1);
      $rolworksheet->getCell('C'.($rolmax+2))->setValue(1);
      $rolworksheet->getCell('C'.($rolmax+3))->setValue(1);
      $rolworksheet->getCell('D'.($rolmax+1))->setValue(1);
      $rolworksheet->getCell('D'.($rolmax+2))->setValue(1);
      $rolworksheet->getCell('D'.($rolmax+3))->setValue(1);
      $rolworksheet->getCell('G'.($rolmax+1))->setValue($_POST[Clarificacion1]);
      $rolworksheet->getCell('G'.($rolmax+2))->setValue($_POST[Clarificacion2]);
      $rolworksheet->getCell('G'.($rolmax+3))->setValue($_POST[Clarificacion3]);
      for ($i=0; $i <3 ; $i++) {
        $rolworksheet->getCell('E'.($rolmax+$i+1))->setValue($_POST['id_solicitud']);
        $rolworksheet->getCell('F'.($rolmax+$i+1))->setValue($_POST['Fecha']);
        $rolworksheet->getCell('H'.($rolmax+$i+1))->setValue($_POST['Marca']);
        $rolworksheet->getCell('I'.($rolmax+$i+1))->setValue($_POST['Modelo']);
        $rolworksheet->getCell('J'.($rolmax+$i+1))->setValue($_POST['Especificacion']);
        $rolworksheet->getCell('K'.($rolmax+$i+1))->setValue($_POST['Control1']);
        $rolworksheet->getCell('L'.($rolmax+$i+1))->setValue($_POST['CantidadPresupuestada']);
        $rolworksheet->getCell('M'.($rolmax+$i+1))->setValue($_POST['CantidadConsumida']);
        $rolworksheet->getCell('N'.($rolmax+$i+1))->setValue($_POST['CantidadCompra']);
        $rolworksheet->getCell('O'.($rolmax+$i+1))->setValue($_POST['Saldo1']);
        $rolworksheet->getCell('P'.($rolmax+$i+1))->setValue($_POST['Observacion1']);
        $rolworksheet->getCell('R'.($rolmax+$i+1))->setValue($_POST['Control2']);
        $rolworksheet->getCell('S'.($rolmax+$i+1))->setValue($_POST['MontoPresupuestado']);
        $rolworksheet->getCell('T'.($rolmax+$i+1))->setValue($_POST['MontoConsumido']);
        $rolworksheet->getCell('U'.($rolmax+$i+1))->setValue($_POST['MontoCompra']);
        $rolworksheet->getCell('V'.($rolmax+$i+1))->setValue($_POST['Saldo2']);
        $rolworksheet->getCell('W'.($rolmax+$i+1))->setValue($_POST['Observacion2']);
      }
      $writer2 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($rolsheet, 'Xlsx');
      $writer2->save('RolesItFO.xlsx');




      $worksheet->getCell('A'.($max+1))->setValue($_POST['id_user1']);
      $worksheet->getCell('A'.($max+2))->setValue($_POST['id_user2']);
      $worksheet->getCell('A'.($max+3))->setValue($_POST['id_user3']);
      $worksheet->getCell('A'.($max+4))->setValue($_POST['id_user4']);
      $worksheet->getCell('A'.($max+5))->setValue($_POST['id_user5']);
      $worksheet->getCell('A'.($max+6))->setValue($_POST['id_user6']);
      $worksheet->getCell('I'.($max+1))->setValue($_POST['Comentario1']);
      $worksheet->getCell('J'.($max+1))->setValue($_POST['estado1']);
      $worksheet->getCell('I'.($max+2))->setValue($_POST['Comentario2']);
      $worksheet->getCell('J'.($max+2))->setValue($_POST['estado2']);
      $worksheet->getCell('I'.($max+3))->setValue($_POST['Comentario3']);
      $worksheet->getCell('J'.($max+3))->setValue($_POST['estado3']);
      $worksheet->getCell('I'.($max+4))->setValue($_POST['Comentario4']);
      $worksheet->getCell('J'.($max+4))->setValue($_POST['estado4']);
      $worksheet->getCell('I'.($max+5))->setValue($_POST['Comentario5']);
      $worksheet->getCell('J'.($max+5))->setValue($_POST['estado5']);
      $worksheet->getCell('I'.($max+6))->setValue($_POST['Comentario6']);
      $worksheet->getCell('J'.($max+6))->setValue($_POST['estado6']);


      $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
      $writer->save('SolicitudHardwareSoftwareFO.xlsx');
      header("Location: /formulario/index.html");
    }
      ?>
      <a title="toyota" href="/formulario/index.html"><img src="/formulario/toyota.png" alt=""></a>
      <br>
      <form class="" action="/formulario/SolicitudAdquisicion.php" method="post">
        <label for="Tipo">Tipo de solicitud</label>
        <select name="Tipo">
         <option>Solicitud de Adquisicion</option>
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
        <label for="Detalle">Detalle del Requerimiento(*)</label>
        <input type="text" id="Detalle" name="Detalle" value="" required>
        <br>
        <label for="Justificacion">Justificacion(*)</label>
        <input type="text" id="Justificacion" name="Justificacion" value="" required>
        <br>
        <label for="Informacion">Informacion de respaldo y anexos</label>
        <input type="text" id="Informacion" name="Informacion" value="">
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
        <br>
        <label for="Fecha">Fecha IT recepcion</label>
        <input type="text" id="Fecha" name="Fecha" value="">
        <br>
        <label for="Clarificacion1">Clarificacion analista IT</label>
        <input type="text" id="Clarificacion1" name="Clarificacion1" value="">
        <br>
        <label for="Clarificacion2">Clarificacion Jefe IT</label>
        <input type="text" id="Clarificacion2" name="Clarificacion2" value="">
        <br>
        <label for="Clarificacion3">Clarificacion gerente IT</label>
        <input type="text" id="Clarificacion3" name="Clarificacion3" value="">
        <br>
        <label for="Marca">Marca</label>
        <input type="text" id="Marca" name="Marca" value="">
        <br>
        <label for="Modelo">Modelo</label>
        <input type="text" id="Modelo" name="Modelo" value="">
        <br>
        <label for="Especificacion">Especificacion tecnica</label>
        <input type="text" id="Especificacion" name="Especificacion" value="">
        <br>
        <label for="Control1">Presupuesto aceptado 1er control</label>
        <input type="text" id="Control1" name="Control1" value="">
        <br>
        <label for="CantidadPresupuestada">Cantidad presupuestada 1er control</label>
        <input type="text" id="CantidadPresupuestada" name="CantidadPresupuestada" value="">
        <br>
        <label for="CantidadConsumida">Cantidad ya consumida 1er control</label>
        <input type="text" id="CantidadConsumida" name="CantidadConsumida" value="">
        <br>
        <label for="CantidadCompra">Cantidad de esta compra 1er control</label>
        <input type="text" id="CantidadCompra" name="CantidadCompra" value="">
        <br>
        <label for="Saldo1">Saldo IT 1er control</label>
        <input type="text" id="Saldo1" name="Saldo1" value="">
        <br>
        <label for="Observacion1">Observacion 1er control</label>
        <input type="text" id="Observacion1" name="Observacion1" value="">
        <br>
        <label for="Aclaracion1">Aclaracion 1er control</label>
        <input type="text" id="Aclaracion1" name="Aclaracion1" value="">
        <br>
        <label for="Control2">Presupuesto aceptado 2do control</label>
        <input type="text" id="Control2" name="Control2" value="">
        <br>
        <label for="MontoPresupuestado">Monto Presupuestodo 2do control</label>
        <input type="text" id="MontoPresupuestado" name="MontoPresupuestado" value="">
        <br>
        <label for="MontoConsumido">Monto consumido 2do control</label>
        <input type="text" id="MontoConsumido" name="MontoConsumido" value="">
        <br>
        <label for="MontoCompra">Monto de esta compra 2do control</label>
        <input type="text" id="MontoCompra" name="MontoCompra" value="">
        <br>
        <label for="Saldo2">Saldo IT 2do control</label>
        <input type="text" id="Saldo2" name="Saldo2" value="">
        <br>
        <label for="Observacion2">Observacion 2do control</label>
        <input type="text" id="Observacion2" name="Observacion2" value="">
        <br>
        <label for="Aclaracion2">Aclaracion 2do control</label>
        <input type="text" id="Aclaracion2" name="Aclaracion2" value="">
        <br>


        <button type="submit" class="btn btn-primary mb-2">Enviar</button>
      </form>

    </body>
  </html>
