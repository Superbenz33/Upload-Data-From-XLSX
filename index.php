<meta charset="utf-8">
<?php

  date_default_timezone_set("Asia/Bangkok");
  require (__DIR__ . '/xlsxreader/XLSXReader.php');

        // entire workbook
        $xlsx = new XLSXReader('#');

        $sheetNames = $xlsx->getSheetNames();

        // loop through worksheets
        foreach ($sheetNames as $sheetName) {
          $sheet = $xlsx->getSheet($sheetName);
          $xlsx_data = $sheet->getData();
          $header_row_xlsx = array_shift($xlsx_data);

          $cnt_col_name = count($header_row_xlsx);

          $row_number = 0;
          foreach ($xlsx_data as $row_xlsx) {
              for ($i = 0; $i < count($row_xlsx); $i++) {
                  $xlsx_field_name = '' . ($i < count($header_row_xlsx) ? $header_row_xlsx[$i] : '');
                  if ($xlsx_field_name === "DoB") {
                      // date value
                      $xlsx_field_value = $row_xlsx[$i];
                  } else {
                      // non-date value
                      $xlsx_field_value = $row_xlsx[$i];
                  }
                  $namedDataArray[$row_number][$xlsx_field_name] = $xlsx_field_value;
              }
              $row_number++;
          }

          // Insert ------------------------------------------------------------

          // column name for insert
          $column_name = '';
          for($c=0;$c < count($header_row_xlsx);$c++){
            $column_name .= $header_row_xlsx[$c].",";
          }
          // column name for insert

          $col_name = substr($column_name, 0, -1);
          foreach ($namedDataArray as $key1 => $value1) {

            $strSQL = "";
            $strSQL1 = "INSERT INTO # ($col_name)";
            $strSQL1 .= " VALUES (";
            $i2 = 0;

            foreach ($value1 as $key2 => $value2) {
              $i2++;
              if($i2 < $cnt_col_name){
                $strSQL1 .= "'".$value2."',";
              }else{
                $strSQL1 .= "'".$value2."'";
              }
            }

            $strSQL1 .= ")";
          }

          // update ------------------------------------------------------------

          foreach ($namedDataArray as $key1 => $value1) {
            $strSQL = "";
            $strSQL1 = "UPDATE # SET";
            $i2 = 0;
            foreach ($value1 as $key2 => $value2) {
              $i2++;
              if($i2 < $cnt_col_name){
                if($key2 != $header_row_xlsx[0] && $key2 != $header_row_xlsx[1]){
                  $strSQL1 .= " $key2 = '".$value2."',";
                }
              }else{
                if($key2 != $header_row_xlsx[0] && $key2 != $header_row_xlsx[1]){
                  $strSQL1 .= " $key2 = '".$value2."'";
                }
              }

            }

            $strSQL1 .= " WHERE $header_row_xlsx[0] = '".$value1["$header_row_xlsx[0]"]."' and $header_row_xlsx[1] = '".$value1["$header_row_xlsx[1]"]."'";
          }

        }
?>
