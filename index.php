<meta charset="utf-8">
<?php

  date_default_timezone_set("Asia/Bangkok");
  require (__DIR__ . '/xlsxreader/XLSXReader.php');

  $Table_name = 'Table3';

  // entire workbook
  $xlsx = new XLSXReader('#');
  /* Get Sheetname = table name */
  $sheetNames = $xlsx->getSheetNames();

  if( in_array($Table_name,$sheetNames) ){

    /* Loop data in sheetname select */
    foreach ($sheetNames as $sheetName) {
      $sheet = $xlsx->getSheet($sheetName);
      $xlsx_data = $sheet->getData();
      $header_row_xlsx = array_shift($xlsx_data);
      $cnt_col_name = count($header_row_xlsx)-1;
      $row_number = 0;

      /* Loop data on all sheet to assoc array */
      foreach ($xlsx_data as $row_xlsx) {
        for ($i = 0; $i < count($row_xlsx); $i++) {
          $xlsx_field_name = '' . ($i < count($header_row_xlsx) ? $header_row_xlsx[$i] : '');
          $xlsx_field_value = $row_xlsx[$i];
          $namedDataArray[$sheetName][$row_number][$xlsx_field_name] = $xlsx_field_value;
        }
        $row_number++;
      }
    }

  }else{
    echo "<p style='color: red;font-size: 18px;'>Can't read you sheetname. Please check you excel file or sheetname ?</p>";
    die();
  }


  /* Section DB Query */
  $fn_db_text = new GenSqlText;
  $strSQL_ins = '';
  $data_req = $namedDataArray[$Table_name];
  $data_res = $fn_db_text->GenQueryText($data_req,$cnt_col_name,$header_row_xlsx,$Table_name);
  print_r($data_res);

  /* End Section DB Query */


  /* 
  ** Class Name : DB Query Insert, Update 
  ** Create by : Benz.Surachai
  */
  class GenSqlText {

    public function GenQueryText($data_req,$cnt_col_name,$header_row_xlsx,$Table_name){
      
      $strSQL_ins = '';

      /* Get Column Name */
      $column_name = '';
      for($c=0;$c < count($header_row_xlsx);$c++) {
        if($header_row_xlsx[$c] != 'DB_type') {
          $column_name .= $header_row_xlsx[$c].",";
        }
      }
      /* End Get Column Name */

      foreach ($data_req as $key1 => $value1) {
        
        if($value1['DB_type'] == 'insert') { /* If Type = insert */

          $strSQL = '';
          $strSQL = "insert into $Table_name (".substr($column_name,0,-1).")";
          $strSQL .= " values (";
          $i2 = 0;

          foreach ($value1 as $key2 => $value2) {
            $i2++;
            if($value2 != 'insert' && $value2 != 'update') {
              if($i2 < $cnt_col_name){
                $strSQL .= '"'.$value2.'",';
              }else{
                $strSQL .= '"'.$value2.'"';
              }
            }
          }

          $strSQL .= "); ";

        } elseif($value1['DB_type'] == 'update') { /* If Type = update */
        
          $strSQL = "";
          $strSQL = "update $Table_name set ";
          $i2 = 0;

          foreach ($value1 as $key2 => $value2) {
            $i2++;
            if($value2 != 'insert' && $value2 != 'update') {
              if($key2 != $header_row_xlsx[0]) {
                if($i2 < $cnt_col_name){
                  $strSQL .= $key2.'='.'"'.$value1[$key2].'", ';
                } else {
                  $strSQL .= $key2.'='.'"'.$value1[$key2].'"';
                }
              }
            }
          }
          /* where column 0 */
          $strSQL .= ' where '.$header_row_xlsx[0].' = "'.$value1[$header_row_xlsx[0]].'"; ';

        } elseif($value1['DB_type'] == 'null' || $value1['DB_type'] == '') { /* If Type = null or empty */
          $strSQL = "";
        } else {
          $strSQL = "";
        }

        $res[] = $strSQL;

      }

      return array_filter($res);
  
    }

  }
?>
