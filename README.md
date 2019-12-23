# php-excel-sheet
phpExcel在Thinkphp5中的应用，可生成多个sheet

使用方法

$excel = new Excel();

$excel->sheet(2)

      ->filename('测试.xlsx')
      
      ->sheetname('工作簿名|工作簿2')
      
      ->title(['标题1','标题2'])
      
      ->header(['李四:aaa:15:align', '张三:test:16:left'])
      
      ->border('thin')
      
      ->body($data)
      
      ->write();
