<?php


namespace api\libs;

use \yii\base\Exception;

/**
 * 通过生成 Excel
 *
 * 说明：支持两种模式：XML、HTML，默认 XML；
 * 两种模式不同：
 *  1、XML：不支持插入图片，支持多工作表；HTML：支持插入图片（直接使用<img>标签，下面有示例），不支持支持多工作表 ！！！
 *  2、自定义样式（看 $this->style 介绍）
 *  3、表格中内容 XML 模式只支持纯字符串或者数字，HTML 模式支持 html 标签
 *
 * 1、简单使用方式：
 *  $headers = ['姓名', '电话'];
 *  $data = [
 *      ['test1', '176'], // 如果是 HTML 模式，['<img src="http://***"/>', '176'] 这样就可以插入图片，或者其他标签，XML 不支持
 *      ['test2', '186']
 *  ];
 *
 *  $excel = new ExcelXml('/tmp/test.xls');
 *
 *  // 如果想多表，重复下面这一部分
 *  $excel->createSheet('表1'); // 名称要唯一
 *  $excel->addRow($headers);
 *
 *  foreach($data as $item){
 *      $excel->addRow($headers);
 *  }
 *
 *  $excel->done();
 *
 *
 *
 * 2、如果需要对单元格进行操作：
 *
 * $fileFullName = Yii::$app->runtimePath.'/excel.xls';
 *
 * $headers = [['value' => '姓名', 'styleId' => 's500'], ['value' => '电话', 'styleId' => 's55'], ['value' => '学习成绩', 'mergeAcross' => 2, 'mergeDown' => 1, 'styleId' => 's55']];
 * $data = [
 * [],
 * ['', ['value' => '超链接', 'styleId' => 's41', 'href' => 'http://www.baidu.com'], '数学', '语文', '英语', '总成绩'],
 * ['小崔', '176', ['value' => 88, 'type' => 1, 'styleId' => 's51'], ['value' => 80, 'type' => 1], ['value' => 78.5, 'type' => 1], ['value' => 0, 'type' => 1, 'formula' => '=RC[-3]+RC[-2]+RC[-1]']],
 * ['flyu', '193', ['value' => 50, 'type' => 1, 'styleId' => 's52'], ['value' => 66, 'type' => 1, 'styleId' => 's53'], ['value' => 70, 'type' => 1, 'styleId' => 's54']],
 * ['aa', '134', ['value' => 35, 'type' => 1, 'styleId' => 's56'], ['value' => 99, 'type' => 1, 'styleId' => 's57'], ['value' => 10, 'type' => 1, 'styleId' => 's58']],
 * ];
 *
 * // XML 新增样式
 * $style = <<<EOF
 * <Style ss:ID="s500" ss:Name="红色字体">
 * <Font ss:Color="#FF0000" ss:Italic="1" ss:Bold="1"/>
 * <Alignment ss:Horizontal="Center"/>
 * </Style>
 *EOF;
 *
 * $excel = new ExcelXml($fileFullName, ['fontName' => 'Yuanti SC Regular', 'style' => $style]);
 *
 * $excel->createSheet('Sheet1', ['column' => [['index' => 1, 'width' => 101.65], ['index' => 5, 'width' => 144.1]]]);
 *
 * $excel->addRow($headers);
 *
 * foreach ($data as $item){
 *  $excel->addRow($item);
 * }
 *
 * $excel->done();
 *
 * 3、对于需要多表同时操作的
 *
 * // 表 1 作为主表和上面一样
 * $headers = ['姓名1', '电话1'];
 * $excel = new Excel($fileFullName);
 * $excel->createSheet('表1');
 * $excel->addRow($headers);
 *
 * $data = [['小崔', '176']];
 * foreach ($data as $item){
 *  $excel->addRow($item);
 *}
 *
 * // 但是下面就不一样了
 * $excel->doneSheet(); // 补充主表未完成的结构（如果先写一张表再写一张不需要这样）
 *
 * $excel2 = new Excel($fileFullName.'_tmp', ['isStart' => false]); // isStart 必须
 * $excel2->createSheet('表2');
 * $excel2->addRow($headers);
 * foreach ($data as $item){
 *  $excel2->addRow($item);
 *}
 *
 * $excel->write($excel2->read()); // 当两张表都完成的时候合并
 * $excel2->del(); // 释放资源
 *
 * // 最终完成
 * $excel->done();
 */
class Excel
{
    // 文档原始格式
    const TYPE_XML = 'XML';
    const TYPE_HTML = 'HTML';

    // 当前文档原始格式，默认 XML
    protected $type = self::TYPE_XML;

    // 当前一共多少个工作表
    private $sheetCount = 0;

    // 打开的文件
    private $handle = null;

    // 文件路径（包括文件名的全路径）
    public $filePath = '';

    // 默认字体
    private $fontName = "宋体";

    // 默认字体大小
    private $fontSize = "12";

    // 默认字体颜色
    private $fontColor = "#000000";

    // 当前列号
    private $currentColNo = 'A';

    // 当前行号
    private $currentRowNo = 1;

    /**
     * 新增样式
     *
     * XML 模式格式如下，ss::ID 不能和默认的重复，具体参数值可以在 excel 里面设置好想要的样式，然后转存成 xml，查看对应的样式就好了；
     * 设置之后就可以通过指定 styleId，进行样式设置了
     *
     *$style = <<<EOF
     *<Style ss:ID="s500" ss:Name="红色字体">
     *<Font ss:Color="#FF0000" ss:Italic="1" ss:Bold="1"/>
     *<Alignment ss:Horizontal="Center"/>
     *</Style>
     *EOF;
     *
     * HTML 模式样式如下，普通的css，设置之后就可以通过指定 styleId，进行样式设置了
     * $style = <<<EOF
     * .s500 {color:#FF0000; text-align:center;}
     *EOF;
     */
    private $style = '';

    // 单元格样式对应表
    private static $cellStyleIdMap = [
        "s41" => "超链接", // 这个只是样式，不能产生效果，需要和 href 属性配合使用
        "s51" => "红色字体",
        "s52" => "黄色背景",
        "s53" => "倾斜",
        "s54" => "加粗",
        "s55" => "居中",
        "s56" => "删除线",
        "s57" => "下划线",
        "s58" => "上标",
        "s59" => "下标",
    ];

    /**
     * 初始化
     *
     * @param string $filePath 包括文件名的全路径
     * @param array $config 全局配置
     *  + string type       模式，默认 XML,可选：HTML
     *  + string fontName   字体名称
     *  + string fontSize   字体大小
     *  + string fontColor  字体颜色
     *  + string style      新增的样式，具体格式看 $this->style
     *  + boolean isStart   是否直接开始生成 excel 文件，默认 true，没有特殊需要不用改
     */
    public function __construct($filePath, $config = [])
    {
        if (!empty($config['type']) && self::TYPE_HTML == strtoupper($config['type'])) $this->type = self::TYPE_HTML;
        if (!empty($config['fontName'])) $this->fontName = $config['fontName'];
        if (!empty($config['fontSize'])) $this->fontSize = $config['fontSize'];
        if (!empty($config['fontColor'])) $this->fontColor = $config['fontColor'];
        if (!empty($config['style'])) $this->style = $config['style'];

        $this->filePath = trim($filePath);
        $this->handle = fopen($this->filePath, "wb+");

        if (!isset($config['isStart']) || true === $config['isStart']) {
            $this->start();
        }
    }

    // 关闭文件
    public function __destruct()
    {
        @fclose($this->handle);
    }

    /**
     * 读取当前正在操作的文件（除非特殊需要，不要这样操作）
     *
     * 注意：因为操作了文件指针，所以不要在写入文件的过程中读取，要不然容易出错 ！！！
     *
     * @param int $size 默认0：读取所有
     * @return string
     */
    public function read($size = 0)
    {
        $ftell = ftell($this->handle);  // 获取当前指针位置
        rewind($this->handle); // 指针指向开头
        $content = fread($this->handle, $size ? $size : $ftell); // 如果不传来大小，默认获取到当前文件的最后面
        fseek($this->handle, $ftell); // 恢复当前文件指针位置

        return $content;
    }

    /**
     * 写入当前正在操作的文件（除非特殊需要，不要这样操作）
     *
     * @param string $content 要写入的数据
     * @return boolean
     */
    public function write($content)
    {
        return fwrite($this->handle, $content);
    }

    /**
     * 删除当前正在操作的文件（除非特殊需要，不要这样操作）
     */
    public function del()
    {
        fclose($this->handle);
        return @unlink($this->filePath);
    }

    // 创建最外层
    private function createRoot()
    {
        fwrite($this->handle, $this->getRootCode());
    }

    // 创建最外层的最后一部分
    private function doneRoot()
    {
        fwrite($this->handle, $this->getRootEndCode());
    }

    // 创建样式
    private function createStyle()
    {
        fwrite($this->handle, $this->getStyleCode());
    }

    /**
     * 创建一个工作表
     *
     * 注意：XML 模式下支持多个工作表，HTML 格式下支持支一个
     *
     * @param string $sheetName 工作表名要唯一，！！！注意：$this->type == HTML 的时候设置无效
     * @param array $ext 扩展功能，如：['column' => [['index' => 1, 'width' => 35]]]
     *      + int   column  当前工作表单元格列设置，['index' => 第几列，从 1 开始, 'width' => 要设置的宽度]
     */
    public function createSheet($sheetName = 'Sheet', $ext = [])
    {
        if (self::TYPE_HTML == $this->type && $this->sheetCount > 0) {
            throw new Exception('HTML模式不支持多工作表');
        }

        // 如果之前没有工作表就算了，如果有的话，先结束之前的
        if ($this->sheetCount > 0) {
            $this->doneSheet();
        }

        // 恢复行数
        $this->currentRowNo = 1;

        $this->sheetCount++;
        $sheetName = !empty($sheetName) && 'Sheet' != $sheetName ? $sheetName : $sheetName . $this->sheetCount;

        fwrite($this->handle, $this->getSheetCode($sheetName));

        if (!empty($ext['column'])) {
            fwrite($this->handle, $this->getColumnCode($ext['column']));
        }
    }

    // 创建工作表的最后一部分
    public function doneSheet()
    {
        fwrite($this->handle, $this->getSheetEndCode());
    }

    /**
     * 添加一行数据
     *
     * @param array $row_data ['第一列', '第二列', [带扩展数据], '第四列']
     *  + 第一列数据,    这样的话使用默认单元格设置
     *  + 第二列数据,
     *  + [    如果需要对单元格特殊操作这样写
     *      + mixed     value           该列的值，如果使用了公式，这个可以写0或者空，会自动计算
     *      + int       type            单元格类型，默认 0：String，1：Number
     *      + int       mergeDown       向下合并指定数量单元格（默认 0）
     *      + int       mergeAcross     向右合并指定数量单元格（默认 0）
     *      + string    href            超链接地址，如果写了这个，styleId 要写 s41
     *      + string    formula         计算公式，如求和：'=RC[-3]+RC[-2]+RC[-1]'
     *          （相对地址，R[如果是负的表示上多少行，正的表示下多少行，不存在表示当前行]C[表示列，和R相同]，如果不知道可以调用 self::positionAbsoluteToRelative() 获取）
     *      + string    styleId         样式id，self::$cellStyleIdMap 中有对照表，如果都不满足，可以自定义，看：$this->style
     *    ],
     *  + 第四行数据
     * @param array $ext 行数据扩展功能，如：['height' => 35]
     *  + int   height  行高
     *
     * 示例：看文件最上面
     */
    public function addRow($row_data, $ext = [])
    {
        $row_str = self::joinRowData($row_data, $ext);
        fwrite($this->handle, $row_str);

        $this->currentRowNo++;
    }

    /**
     * 添加多行数据
     *
     * 说明：使用这个比行数据减少文件写入次数，但是如果行数太多可能造成内存占用大的问题，所以看需求，不要一次太多
     *
     * @param array $data [[第一行数据（和 self::addRow() 相同）], [第二行数据]……]
     */
    public function batch($data)
    {
        $row_str = '';
        foreach($data as $k => $row){
            $row_str .= self::joinRowData($row);
            $this->currentRowNo++;

            // 500 行插入一次
            if(0 == $k % 500){
                fwrite($this->handle, $row_str);
                $row_str = '';
            }
        }

        // 不足 500 的
        fwrite($this->handle, $row_str);
    }

    /**
     * 拼接行数据（参数和 self::addRow() 相同）
     *
     * @param $row_data
     * @param array $ext
     * @return string
     */
    private function joinRowData($row_data, $ext = [])
    {
        $rowStr = $this->getRowCode($ext);

        foreach ($row_data as $index => $item) {
            if (0 == $index) {
                $this->currentColNo = 'A';
            } else {
                // 查看上一个是否有合并单元格
                if (!empty($row_data[$index - 1]['mergeAcross'])) {
                    // 先把上个单元格合并的计算了
                    $this->currentColNo = chr(ord($this->currentColNo) + (int)$row_data[$index - 1]['mergeAcross']);
                }

                // 当前的加上
                $this->currentColNo++;
            }

            $rowStr .= $this->getCellCode($item);
        }

        $rowStr .= $this->getRowEndCode();

        return $rowStr;
    }

    public function start()
    {
        $this->createRoot();
        $this->createStyle();
    }

    // 处理最后的部分
    public function done()
    {
        $this->doneSheet();
        $this->doneRoot();
    }

    // 获取最外层信息
    private function getRootCode()
    {
        if (self::TYPE_XML == $this->type) {
            $str = <<<EOF
<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
EOF;
        } else {
            $str = <<<EOF
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:x="urn:schemas-microsoft-com:office:excel"
      xmlns="http://www.w3.org/TR/REC-html40">
    <head>
        <meta http-equiv="Content-type" content="text/html;charset=UTF-8" />
        <!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Sheet1</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
EOF;
        }

        return $str;
    }

    // 获取最外层结尾信息
    private function getRootEndCode()
    {
        if (self::TYPE_XML == $this->type) {
            $str = '</Workbook>';
        } else {
            $str = '</html>';
        }
        return $str;
    }

    // 获取样式信息（仅保存了默认的几种样式）
    private function getStyleCode()
    {
        if (self::TYPE_XML == $this->type) {
            //<Style ss:ID="s51" ss:Name="红色字体" ss:Parent="s41"> // 可以指定父级
            $str = <<<EOF
    <Styles>
        <Style ss:ID="Default" ss:Name="Normal">
            <Alignment/>
            <Borders/>
            <Font ss:FontName="{$this->fontName}" x:CharSet="134" ss:Size="{$this->fontSize}" ss:Color="{$this->fontColor}"/>
            <Interior/>
            <NumberFormat/>
            <Protection/>
        </Style>
        <Style ss:ID="s40" ss:Name="默认" ss:Parent="Default">
        </Style>
        <Style ss:ID="s41" ss:Name="超链接">
            <Font ss:FontName="宋体" x:CharSet="0" ss:Size="11" ss:Color="#0000FF" ss:Underline="Single"/>
        </Style>
        <Style ss:ID="s51" ss:Name="红色字体">
            <Font ss:Color="#FF0000"/>
        </Style>
        <Style ss:ID="s52" ss:Name="黄色背景">
            <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
        </Style>
        <Style ss:ID="s53" ss::Name="倾斜">
            <Font ss:Italic="1"/>
        </Style>
        <Style ss:ID="s54" ss::Name="加粗">
            <Font ss:Bold="1"/>
        </Style>
        <Style ss:ID="s55" ss::Name="适中">
            <Alignment ss:Horizontal="Center"/>
        </Style>
        <Style ss:ID="s56" ss::Name="删除线">
            <Font ss:StrikeThrough="1"/>
        </Style>
        <Style ss:ID="s57" ss::Name="下划线">
            <Font ss:Underline="Single"/>
        </Style>
        <Style ss:ID="s58" ss::Name="上标">
            <Font ss:VerticalAlign="Superscript"/>
        </Style>
        <Style ss:ID="s59" ss::Name="下标">
            <Font ss:VerticalAlign="Subscript"/>
        </Style>
        {$this->style}
    </Styles>
EOF;
        } else {
            $str = <<<EOF
    <style>
    tr
        {mso-height-source:auto;
        mso-ruby-visibility:none;}
    col
        {mso-width-source:auto;
        mso-ruby-visibility:none;}
    br
        {mso-data-placement:same-cell;}
    
    .Default
	{mso-number-format:"General";
	text-align:general;
	vertical-align:middle;
	white-space:nowrap;
	mso-rotate:0;
	color:{$this->fontColor};
	font-size:{$this->fontSize}pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:{$this->fontName};
	mso-font-charset:134;
	border:none;
	mso-protection:locked visible;
	mso-style-name:"常规";
	mso-style-id:0;} 
	
    td
	{mso-style-parent:Default;
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	mso-number-format:"General";
	text-align:general;
	vertical-align:middle;
	white-space:nowrap;
	mso-rotate:0;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	border:none;
	mso-protection:locked visible;}
	
	.s40 /*默认*/
	{mso-style-parent:Default;}
	
	.s41 /*超链接（配合html标签a使用）*/
	{mso-style-parent:Default;
	color:#0000FF;
	font-size:11.0pt;
	text-decoration:underline;
	text-underline-style:single;
	mso-font-charset:0;}
	
	.s51 /*红色字体*/
	{mso-style-parent:Default;
	color:#FF0000;}
	
	.s52 /*黄色背景*/
	{mso-style-parent:Default;
	color:#FFFF00;}
	
	.s53 /*倾斜*/
	{mso-style-parent:Default;
	font-style:italic}
	
	.s54 /*加粗*/
	{mso-style-parent:Default;
	font-weight:700;}
	
    .s55 /*适中*/
	{mso-style-parent:Default;
	text-align:center;}
	
	.s56 /*删除线（配合html标签del使用）*/
	{mso-style-parent:Default;}
	
	.s57 /*下划线*/
	{mso-style-parent:Default;
	text-decoration:underline;
	text-underline-style:single;}
	
	.s58 /*上标（配合html标签sup使用）*/
	{mso-style-parent:Default;}
	
	.s59 /*下标（配合html标签sub使用）*/
	{mso-style-parent:Default;}
	
    {$this->style}
     -->  </style>
    </head>
EOF;
        }

        return $str;
    }

    /**
     * 获取工作表信息
     *
     * @param string $sheetName 不能重复，注意：$this->type == HTML 的时候设置无效
     * @return string
     */
    private function getSheetCode($sheetName)
    {
        if (self::TYPE_XML == $this->type) {
            $str = <<<EOF
    <Worksheet ss:Name="{$sheetName}">
        <Table x:FullColumns="1" x:FullRows="1" ss:DefaultColumnWidth="53" ss:DefaultRowHeight="15" ss:StyleID="s40">
EOF;
        } else {
            $str = <<<EOF
    <body link="blue" vlink="purple" class="s40">
        <table width="420.55" border="0" cellpadding="0" cellspacing="0" style='width:420.55pt;border-collapse:collapse;table-layout:fixed;'>
EOF;
        }

        return $str;
    }

    // 获取工作表结尾信息
    private function getSheetEndCode()
    {
        if (self::TYPE_XML == $this->type) {
            $str = '</Table></Worksheet>';
        } else {
            $str = '</table></body>';
        }

        return $str;
    }

    /**
     * 获取列信息
     *
     * @param array $attribute 如：[['index' => 第几列，从 1 开始, 'width' => 要设置的宽度]]
     * @return string
     */
    private function getColumnCode($attribute = [])
    {
        $str = '';
        if (self::TYPE_XML == $this->type) {
            foreach ($attribute as $item) {
                $str .= '<Column ss:Index="' . $item['index'] . '" ss:StyleID="s49" ss:AutoFitWidth="0" ss:Width="' . $item['width'] . '"/>';
            }
        } else {
            // 先找出最大设置的列
            $maxIndex = 0;
            $_attribute = [];
            foreach ($attribute as $item) {
                if ($item['index'] > $maxIndex) {
                    $maxIndex = $item['index'];
                }
                $_attribute[$item['index']] = $item['width'];
            }
            $defaultWidth = 47.80;

            for ($index = 1; $index <= $maxIndex; $index++) {
                if (isset($_attribute[$index])) {
                    $str .= '<col width="' . $_attribute[$index] . '" class="s40"/>';
                } else {
                    $str .= '<col width="' . $defaultWidth . '" class="s40"/>';
                }
            }
        }

        return $str;
    }

    /**
     * 获取行信息
     *
     * @param array $attribute 如：['height' => 35]
     * @return string
     */
    private function getRowCode($attribute = [])
    {
        if (self::TYPE_XML == $this->type) {
            $row = '<Row ';
            if (!empty($attribute['height']) && is_numeric($attribute['height'])) {
                $row .= 'ss:Height="' . $attribute['height'] . '">';
            } else {
                $row .= '>';
            }
        } else {
            $row = '<tr ';
            if (!empty($attribute['height']) && is_numeric($attribute['height'])) {
                $row .= 'height="' . $attribute['height'] . '" style="' . $attribute['height'] . 'pt">';
            } else {
                $row .= '>';
            }
        }

        return $row;
    }

    // 获取行结尾信息
    private function getRowEndCode()
    {
        if (self::TYPE_XML == $this->type) {
            $str = '</Row>';
        } else {
            $str = '</tr>';
        }

        return $str;
    }

    // 获取单元格信息（结构看 self::addRow() 中的 $data）
    private function getCellCode($cellData)
    {
        if (self::TYPE_XML == $this->type) {
            $cell = '<Cell ';

            if (is_array($cellData)) {
                if (!empty($cellData['styleId'])) $cell .= ' ss:StyleID="' . $cellData['styleId'] . '" ';
                if (!empty($cellData['mergeDown'])) $cell .= ' ss:MergeDown="' . $cellData['mergeDown'] . '" ';
                if (!empty($cellData['mergeAcross'])) $cell .= ' ss:MergeAcross="' . $cellData['mergeAcross'] . '" ';
                if (!empty($cellData['href'])) $cell .= ' ss:HRef="' . $cellData['href'] . '" ';
                $type = !empty($cellData['type']) ? 'Number' : 'String';

                // 如果设置了公式，类型直接强制改成 num
                if (!empty($cellData['formula'])) {
                    $cell .= ' ss:Formula="' . $cellData['formula'] . '" ';
                    $type = 'x:num';
                }

                $value = isset($cellData['value']) ? $cellData['value'] : '';

                $cell .= '><Data ss:Type="' . $type . '">' . $value . '</Data></Cell>';
            } else {
                $cell .= '><Data ss:Type="String">' . $cellData . '</Data></Cell>';
            }
        } else {
            $cell = '<td ';
            $styleId = (!empty($cellData['styleId']) ? $cellData['styleId'] : 's40');
            $cell .= ' class="' . $styleId . '" ';

            if (is_array($cellData)) {
                if (!empty($cellData['mergeDown'])) $cell .= ' rowspan="' . ($cellData['mergeDown'] + 1) . '" ';
                if (!empty($cellData['mergeAcross'])) $cell .= ' colspan="' . ($cellData['mergeAcross'] + 1) . '" ';

                $type = !empty($cellData['type']) ? 'x:num' : 'x:str';

                // 如果设置了公式，类型直接强制改成 num
                if (!empty($cellData['formula'])) {
                    // html 中公式要用绝对地址，做个转换
                    $formula = self::positionRelativeToAbsolute($cellData['formula'], $this->currentColNo, $this->currentRowNo);
                    $cell .= ' x:fmla="' . $formula . '" ';
                    $type = 'x:num';
                }

                $cell .= ' ' . $type . ' ';
                $cell .= '>';

                $value = isset($cellData['value']) ? $cellData['value'] : '';

                if (!empty($cellData['href'])) {
                    $value = '<a href="' . $cellData['href'] . '" target="_parent">' . $value . '</a>';
                } else if ('s56' == $styleId) { // 删除线
                    $value = '<del>' . $value . '</del>';
                } else if ('s58' == $styleId) { // 上标
                    $value = '<sup>' . $value . '</sup>';
                } else if ('s59' == $styleId) { // 下标
                    $value = '<sub>' . $value . '</sub>';
                }

                $cell .= $value . '</td>';
            } else {
                $cell .= '>' . $cellData . '</td>';
            }
        }

        return $cell;
    }

    /**
     * Excel相对位置转成绝对位置
     *
     * 注意：可以用这个和 self::positionAbsoluteToRelative() 做相互验证
     *
     * @param string $str 格式：=(R[2]C[-2]+RC[-2])/R[-1]C[2]（R:row，C:col，方框同行为空，下一行或者右边为正，上一行或者左边为负）
     * @param string $current_col_no 当前第几列
     * @param int $current_row_no 当前第几行
     *
     * @return string 如果当前位置是F4，返回：=(D6+D4)/H3（excel公式）
     */
    public static function positionRelativeToAbsolute($str, $current_col_no = 'A', $current_row_no = 1)
    {
        $tmp = '';

        // 重组 str
        $new_str = '';
        $arr = preg_split("//", $str, -1);
        foreach ($arr as $k => $item) {
            if ('R' == $item && '[' != $arr[$k + 1] && 'C' == $arr[$k + 1]) {
                $item .= '[0]';
            } else if ('C' == $item && '[' != $arr[$k + 1] && ('R' == $arr[$k - 1] || ']' == $arr[$k - 1])) {
                $item .= '[0]';
            }
            $new_str .= $item;
        }

        // 匹配出所有的坐标
        $match_count = preg_match_all('/([RC])\[(.*?)\]/', $new_str, $result);
        if (empty($match_count)) {
            return '';
        }

        // 用相同的正则做一个模板
        $template_arr = preg_split("/([RC])\[(.*?)\]/", $new_str, -1);

        // 替换成绝对地址
        $empty_index = 0;
        $template_arr_count = count($template_arr);
        foreach ($template_arr as $index => $item) {
            if ('' != $item) {
                $tmp .= $item;
                continue;
            }

            // 空表示要插入的部分
            if (0 == $index || $template_arr_count - 1 == $index) continue;  // 第一个和最后一个空格过滤掉

            $_current_col_no = $current_col_no;
            $_current_row_no = $current_row_no;

            $c_relative = $result[2][$empty_index * 2 + 1];
            $r_relative = $result[2][$empty_index * 2];

            $_current_col_no_ascall = ord($_current_col_no);
            $_current_col_no_ascall += (int)$c_relative;
            $_current_col_no = chr($_current_col_no_ascall);

            $_current_row_no += (int)$r_relative;

            $tmp .= $_current_col_no . $_current_row_no;

            $empty_index++;
        }

        return $tmp;
    }

    /**
     * Excel绝对位置转成相对位置
     *
     *  注意：可以用这个和 self::positionRelativeToAbsolute() 做相互验证
     *
     * @param string $str 格式：=(D6+D4)/H3（excel公式）
     * @param string $current_col_no 当前第几列
     * @param int $current_row_no 当前第几行
     *
     * @return string 如果当前位置是F4，返回：=(R[2]C[-2]+RC[-2])/R[-1]C[2] （R:row，C:col，方框同行为空，下一行或者右边为正，上一行或者左边为负）
     */
    public static function positionAbsoluteToRelative($str, $current_col_no = 'A', $current_row_no = 1)
    {
        $tmp = '';

        // 匹配出所有的坐标
        $match_count = preg_match_all('/(\w+)(\d+)/', $str, $result);
        if (empty($match_count)) {
            return '';
        }

        // 用相同的正则做一个模板
        $template_arr = preg_split("/(\w+)(\d+)/", $str, -1);

        // 替换成绝对地址
        $empty_index = 0;
        foreach ($template_arr as $index => $item) {
            if ('' == $item) {
                continue;
            }

            $tmp .= $item;

            $_current_col_no = $current_col_no;
            $_current_row_no = $current_row_no;

            $c_absolute = $result[1][$empty_index];
            $r_absolute = $result[2][$empty_index];

            $_target_col_no_ascall = ord($c_absolute);

            $_current_col_no_ascall = ord($_current_col_no);

            $c_relative_col_no = $_target_col_no_ascall - $_current_col_no_ascall;
            $r_relative_row_no = $r_absolute - $_current_row_no;

            $tmp .= 'R' . (0 != $r_relative_row_no ? "[$r_relative_row_no]" : '');
            $tmp .= 'C' . (0 != $c_relative_col_no ? "[$c_relative_col_no]" : '');

            $empty_index++;
        }

        return $tmp;
    }
}
