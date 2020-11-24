# excel
PHP 通过生成 XML 或者 HTML 从而转成 EXCEl，解决了其他生成 EXCEL 的扩展内存过大的问题，但是 XML 模式支持多表格不支持图片，HTML 模式支持 HTML 标签（也就是支持图片）但是不支持多表格，其他没什么区别

## 说明：支持两种模式：XML、HTML，默认 XML，两种模式不同：
1. XML：不支持插入图片，支持多工作表；HTML：支持插入图片（直接使用<img>标签，下面有示例），不支持支持多工作表 ！！！
2. 自定义样式（看 $this->style 介绍）
3. 表格中内容 XML 模式只支持纯字符串或者数字，HTML 模式支持 html 标签

## 基本功能
- 支持 XML、HTML 两种源码模式
- 使用默认配置简单使用
- 自定义全局样式配置
- 单元格（支持：单元格类型、合并单元格、超链、公式、默认可选的单元格样式、自定义单元格样式）
- 行（支持：自定义行高）


## 简单使用方式：

```php
 $headers = ['姓名', '电话'];
 $data = [
     ['test1', '176'], // 如果是 HTML 模式，['<img src="http://***"/>', '176'] 这样就可以插入图片，或者其他标签，XML 不支持
     ['test2', '186']
 ];

 $excel = new Excel('/tmp/test.xls');

 // 如果想多表，重复下面这一部分
 $excel->createSheet('表1'); // 名称要唯一
 $excel->addRow($headers);

 // 如果批量这样（效率更高）
 $excel->batch($data);
 
 /*
 // 如果单行插入这样
 foreach($data as $item){
     $excel->addRow($headers);
 }
 */

 $excel->done();
```

## 如果需要对单元格进行操作：

```php
 $fileFullName = Yii::$app->runtimePath.'/excel.xls';

 $headers = [['value' => '姓名', 'styleId' => 's500'], ['value' => '电话', 'styleId' => 's55'], ['value' => '学习成绩', 'mergeAcross' => 2, 'mergeDown' => 1, 'styleId' => 's55']];
 $data = [
 [],
 ['', ['value' => '超链接', 'styleId' => 's41', 'href' => 'http://www.baidu.com'], '数学', '语文', '英语', '总成绩'],
 ['小崔', '176', ['value' => 88, 'type' => 1, 'styleId' => 's51'], ['value' => 80, 'type' => 1], ['value' => 78.5, 'type' => 1], ['value' => 0, 'type' => 1, 'formula' => '=RC[-3]+RC[-2]+RC[-1]']],
 ['flyu', '193', ['value' => 50, 'type' => 1, 'styleId' => 's52'], ['value' => 66, 'type' => 1, 'styleId' => 's53'], ['value' => 70, 'type' => 1, 'styleId' => 's54']],
 ['aa', '134', ['value' => 35, 'type' => 1, 'styleId' => 's56'], ['value' => 99, 'type' => 1, 'styleId' => 's57'], ['value' => 10, 'type' => 1, 'styleId' => 's58']],
 ];

 // XML 新增样式
 $style = <<<EOF
 <Style ss:ID="s500" ss:Name="红色字体">
 <Font ss:Color="#FF0000" ss:Italic="1" ss:Bold="1"/>
 <Alignment ss:Horizontal="Center"/>
 </Style>
EOF;

 $excel = new Excel($fileFullName, ['fontName' => 'Yuanti SC Regular', 'style' => $style]);

 $excel->createSheet('Sheet1', ['column' => [['index' => 1, 'width' => 101.65], ['index' => 5, 'width' => 144.1]]]);

 $excel->addRow($headers);

 foreach ($data as $item){
  	$excel->addRow($item);
 }

 $excel->done();
```

## 对于需要多表同时操作的

```php
 // 表 1 作为主表和上面一样
 $headers = ['姓名1', '电话1'];
 $excel = new Excel($fileFullName);
 $excel->createSheet('表1');
 $excel->addRow($headers);

 $data = [['小崔', '176']];
 foreach ($data as $item){
  	$excel->addRow($item);
 }

 // 但是下面就不一样了
 $excel->doneSheet(); // 补充主表未完成的结构（如果先写一张表再写一张不需要这样）

 $excel2 = new Excel($fileFullName.'_tmp', ['isStart' => false]); // isStart 必须
 $excel2->createSheet('表2');
 $excel2->addRow($headers);
 foreach ($data as $item){
  	$excel2->addRow($item);
 }

 $excel->write($excel2->read()); // 当两张表都完成的时候合并
 $excel2->del(); // 释放资源

 // 最终完成
 $excel->done();
```

