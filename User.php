<!DOCTYPE html>
<html lang="en" dir="ltr">
  <head>
    <meta charset="utf-8">
      <title>Cargar usuario</title>
      <link rel="stylesheet"  href="/formulario/css/prueba.css">
      <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
  </head>
  <body>
    <?php require 'class/vendor/autoload.php';
      use PhpOffice\PhpSpreadsheet\Spreadsheet;
      use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
      $usersheet= \PhpOffice\PhpSpreadsheet\IOFactory::load('UserFO.xlsx');
      $userworksheet=$usersheet->setActiveSheetIndexByName('UserFO');
      $usermax=$userworksheet->getHighestRow();
      $id=$userworksheet->getCell('A'.$usermax)->getValue();
      if ($_POST) {
        $userworksheet->getCell('A'.($usermax+1))->setValue($_POST['id_user']);
        $userworksheet->getCell('B'.($usermax+1))->setValue($_POST['User']);
        $userworksheet->getCell('C'.($usermax+1))->setValue($_POST['Pass']);
        $userworksheet->getCell('E'.($usermax+1))->setValue($_POST['Id_role']);
        $userworksheet->getCell('D'.($usermax+1))->setValue($_POST['UserExtend']);
        $userworksheet->getCell('F'.($usermax+1))->setValue($_POST['Name_role']);
        $userworksheet->getCell('G'.($usermax+1))->setValue($_POST['Name_user']);
        $userworksheet->getCell('H'.($usermax+1))->setValue($_POST['Docket']);

        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($usersheet, 'Xlsx');
      $writer->save('UserFO.xlsx');
      header("Location: /formulario/index.html");
      }

     ?>
     <a title="toyota" href="/formulario/index.html"><img src="/formulario/toyota.png" alt=""></a>
     <br>
      <form class="" action="/formulario/User.php" method="post">
        <label for="id_user">Id_user</label>
     <select class="" name="id_user">
            <option value="<?php echo  $id+1; ?>"><?php echo $id+1; ?></option>
          </select>
          <br>
     <label for="User">User</label>
     <input type="text" id="User" name="User" value="">
     <br>
     <label for="Pass">Password</label>
     <input type="text" id="Pass" name="Pass" value="">
     <br>
     <label for="Id_role">Id_role</label>
     <input type="text" id="Id_role" name="Id_role" value="">
     <br>
     <label for="UserExtend">User_extend_name</label>
     <input type="text" id="UserExtend" name="UserExtend" value="">
     <br>
     <label for="Name_role">name_role</label>
     <input type="text" id="Name_role" name="Name_role" value="">
     <br>
     <label for="Name_user">Name_user</label>
     <input type="text" id="Name_user" name="Name_user" value="">
     <br>
     <label for="Docket">Docket</label>
     <input type="text" id="Docket" name="Docket" value="">
     <br>
     <button type="submit" class="btn btn-primary mb-2">Enviar</button>
   </form>

  </body>
</html>
