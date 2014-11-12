<?php
/*
 * @license MIT License
 * Bug fixes by smiffy6969 as per issues raised on github https://github.com/SystemDevil/PHP_XLSXWriter_plus
 * */

if (!class_exists('ZipArchive')) { throw new Exception('ZipArchive not found'); }

Class XLSXWriter
{
	//------------------------------------------------------------------
	protected $author ='Doc Author';

	protected $defaultFontName = 'Calibri';
	protected $defaultFontSize = 11;
	protected $defaultWrapText = false;
	protected $defaultVerticalAlign = 'top';
	protected $defaultHorizontalAlign = 'left';
	protected $defaultStartRow = 0;
	protected $defaultStartCol = 0;

	protected $defaultStyle = array();

	protected $fontsCount = 1; //1 font must be in structure
	protected $fontSize = 8;
	protected $fontColor = '';
	protected $fontStyles = '';
	protected $fontName = '';
	protected $fontId = 0; //font counting from index - 0, means 0,1 - 2 elements

	protected $bordersCount = 1; //1 border must be in structure
	protected $bordersStyle = '';
	protected $bordersColor = '';
	protected $borderId = 0; //borders counting from index - 0, means 0,1 - 2 elements

	protected $fillsCount = 2; //2 fills must be in structure
	protected $fillColor = '';
	protected $fillId = 1; //fill counting from index - 0, means 0,1 - 2 elements

	protected $stylesCount = 1;//1 style must be in structure

	protected $sheets_meta = array();
	protected $shared_strings = array();//unique set
	protected $shared_string_count = 0;//count of non-unique references to the unique set
	protected $temp_files = array();

	public function __construct(){}
	public function setAuthor($author='') { $this->author=$author; }
	public function setFontName($defaultFontName) { $this->defaultFontName=$defaultFontName; }
	public function setFontSize($defaultFontSize) { $this->defaultFontSize=$defaultFontSize; }
	public function setWrapText($defaultWrapText) { $this->defaultWrapText=$defaultWrapText; }
	public function setVerticalAlign($defaultVerticalAlign) { $this->defaultVerticalAlign=$defaultVerticalAlign; }
	public function setHorizontalAlign($defaultHorizontalAlign) { $this->defaultHorizontalAlign=$defaultHorizontalAlign; }
	private function setStyle($defaultStyle) { $this->defaultStyle=$defaultStyle; }
	public function setStartRow($defaultStartRow) { $this->defaultStartRow=($defaultStartRow > 0) ? ((int)$defaultStartRow - 1) : 0; }
	public function setStartCol($defaultStartCol) { $this->defaultStartCol=($defaultStartCol > 0) ? ((int)$defaultStartCol - 1) : 0; }

	public function __destruct()
	{
		if (!empty($this->temp_files)) {
			foreach($this->temp_files as $temp_file) {
				@unlink($temp_file);
			}
		}
	}
	
	protected function tempFilename()
	{
		$filename = tempnam("/tmp", "xlsx_writer_");
		$this->temp_files[] = $filename;
		return $filename;
	}

	public function writeToStdOut()
	{
		$temp_file = $this->tempFilename();
		self::writeToFile($temp_file);
		readfile($temp_file);
	}

	public function writeToString()
	{
		$temp_file = $this->tempFilename();
		self::writeToFile($temp_file);
		$string = file_get_contents($temp_file);
		return $string;
	}

	public function writeToFile($filename)
	{
		@unlink($filename);//if the zip already exists, overwrite it
		$zip = new ZipArchive();
		if (empty($this->sheets_meta))                  { self::log("Error in ".__CLASS__."::".__FUNCTION__.", no worksheets defined."); return; }
		if (!$zip->open($filename, ZipArchive::CREATE)) { self::log("Error in ".__CLASS__."::".__FUNCTION__.", unable to create zip."); return; }
		
		$zip->addEmptyDir("docProps/");
		$zip->addFromString("docProps/app.xml" , self::buildAppXML() );
		$zip->addFromString("docProps/core.xml", self::buildCoreXML());

		$zip->addEmptyDir("_rels/");
		$zip->addFromString("_rels/.rels", self::buildRelationshipsXML());

		$zip->addEmptyDir("xl/worksheets/");
		foreach($this->sheets_meta as $sheet_meta) {
			$zip->addFile($sheet_meta['filename'], "xl/worksheets/".$sheet_meta['xmlname'] );
		}
		if (!empty($this->shared_strings)) {
			$zip->addFile($this->writeSharedStringsXML(), "xl/sharedStrings.xml" );  //$zip->addFromString("xl/sharedStrings.xml",     self::buildSharedStringsXML() );
		}
		$zip->addFromString("xl/workbook.xml"         , self::buildWorkbookXML() );
		$zip->addFile($this->writeStylesXML(), "xl/styles.xml" );  //$zip->addFromString("xl/styles.xml"           , self::buildStylesXML() );
		$zip->addFromString("[Content_Types].xml"     , self::buildContentTypesXML() );

		$zip->addEmptyDir("xl/_rels/");
		$zip->addFromString("xl/_rels/workbook.xml.rels", self::buildWorkbookRelsXML() );
		$zip->close();
	}

	
	public function writeSheet(array $data, $sheet_name='', array $header_types=array(), array $styles=array())
	{
		for ($i = 0; $i < count($styles); $i++) {
			$styles[$i] += array('sheet' => $sheet_name);
		}
		$this->setStyle(array_merge((array)$this->defaultStyle, (array)$styles));

		$data = empty($data) ? array( array('') ) : $data;

		$sheet_filename = $this->tempFilename();
		$sheet_default = 'Sheet'.(count($this->sheets_meta)+1);
		$sheet_name = !empty($sheet_name) ? $sheet_name : $sheet_default;
		$this->sheets_meta[] = array('filename'=>$sheet_filename, 'sheetname'=>$sheet_name ,'xmlname'=>strtolower($sheet_default).".xml" );

		$header_offset = empty($header_types) ? 0 : $this->defaultStartRow + 1;
		$row_count = count($data) + $header_offset;
		$column_count = count($data[self::array_first_key($data)]);
		$max_cell = self::xlsCell( $row_count-1, $column_count-1 );

		$tabselected = count($this->sheets_meta)==1 ? 'true' : 'false';//only first sheet is selected
		$cell_formats_arr = empty($header_types) ? array_fill(0, $column_count, 'string') : array_values($header_types);
		$header_row = empty($header_types) ? array() : array_keys($header_types);

		$fd = fopen($sheet_filename, "w+");
		if ($fd===false) { self::log("write failed in ".__CLASS__."::".__FUNCTION__."."); return; }
		
		fwrite($fd,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");
		fwrite($fd,'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
		fwrite($fd,    '<sheetPr filterMode="false">');
		fwrite($fd,        '<pageSetUpPr fitToPage="false"/>');
		fwrite($fd,    '</sheetPr>');
		fwrite($fd,    '<dimension ref="A1:'.$max_cell.'"/>');
		fwrite($fd,    '<sheetViews>');
		fwrite($fd,        '<sheetView colorId="64" defaultGridColor="true" rightToLeft="false" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" tabSelected="'.$tabselected.'" topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">');
		fwrite($fd,            '<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>');
		fwrite($fd,        '</sheetView>');
		fwrite($fd,    '</sheetViews>');
		fwrite($fd,    '<cols>');
		fwrite($fd,        '<col collapsed="false" hidden="false" max="1025" min="1" style="0" width="11.5"/>');
		fwrite($fd,    '</cols>');
		fwrite($fd,    '<sheetData>');
		if (!empty($header_row))
		{
			fwrite($fd, '<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="'.($this->defaultStartRow + 1).'">');
			foreach($header_row as $k=>$v)
			{
				$this->writeCell($fd, $this->defaultStartRow + 0, $this->defaultStartCol + $k, $v, $sheet_name);
			}
			fwrite($fd, '</row>');
		}
		foreach($data as $i=>$row)
		{
			fwrite($fd, '<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="'.($i+$header_offset+1).'">');
			foreach($row as $k=>$v)
			{
				$this->writeCell($fd, $i+$header_offset, $this->defaultStartCol + $k, $v, $sheet_name);
			}
			fwrite($fd, '</row>');
		}
		fwrite($fd,    '</sheetData>');
		fwrite($fd,    '<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>');
		fwrite($fd,    '<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>');
		fwrite($fd,    '<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>');
		fwrite($fd,    '<headerFooter differentFirst="false" differentOddEven="false">');
		fwrite($fd,        '<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>');
		fwrite($fd,        '<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>');
		fwrite($fd,    '</headerFooter>');
		fwrite($fd,'</worksheet>');
		fclose($fd);
	}

	protected function writeCell($fd, $row_number, $column_number, $value, $sheet_name)
	{
		$cell = self::xlsCell($row_number, $column_number);
		$s = '0';
		if ($this->defaultStyle) {
			foreach ($this->defaultStyle as $key => $style) {
				if (isset($style['sheet'])) {
					if ($style['sheet'] == $sheet_name) {
						if (isset($style['allfilleddata'])) {
							$s = $key + 1;
						} else {
							if (isset($style['columns'])) {
								if (is_array($style['columns'])) {
									if (in_array($column_number, $style['columns'])) $s = $key + 1;
								} else {
									if ($column_number == $style['columns']) $s = $key + 1;
								}
							} elseif (isset($style['rows'])) {
								if (is_array($style['rows'])) {
									if (in_array($row_number, $style['rows'])) $s = $key + 1;
								} else {
									if ($row_number == $style['rows']) $s = $key + 1;
								}
							} elseif (isset($style['cells'])) {
								if (is_array($style['cells'])) {
									if (in_array($cell, $style['cells'])) $s = $key + 1;
								} else {
									if ($cell == $style['cells']) $s = $key + 1;
								}
							}
						}
					}
				}
			}
		}
		if (is_numeric($value)) {
			fwrite($fd,'<c r="'.$cell.'" s="'.$s.'" t="n"><v>'.($value*1).'</v></c>');//int,float, etc
		} else if ($value=='date') {
			fwrite($fd,'<c r="'.$cell.'" s="'.$s.'" t="n"><v>'.intval(self::convert_date_time($value)).'</v></c>');
		} else if ($value=='datetime') {
			fwrite($fd,'<c r="'.$cell.'" s="'.$s.'" t="n"><v>'.self::convert_date_time($value).'</v></c>');
		} else if ($value==''){
			fwrite($fd,'<c r="'.$cell.'" s="'.$s.'"/>');
		} else if ($value{0}=='='){
			fwrite($fd,'<c r="'.$cell.'" s="'.$s.'" t="s"><f>'.self::xmlspecialchars($value).'</f></c>');
		} else if ($value!==''){
			fwrite($fd,'<c r="'.$cell.'" s="'.$s.'" t="s"><v>'.self::xmlspecialchars($this->setSharedString($value)).'</v></c>');
		}
	}

	protected function writeStylesXML()
	{
		$tempfile = $this->tempFilename();
		$fd = fopen($tempfile, "w+");
		if ($fd===false) { self::log("write failed in ".__CLASS__."::".__FUNCTION__."."); return; }
		fwrite($fd, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");
		fwrite($fd, '<styleSheet xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
		if ($this->defaultStyle) {
			foreach ($this->defaultStyle as $style) {
				if (isset($style['sheet'])) {
					if (isset($style['font'])) $this->fontsCount++;
				}
			}
		}
		fwrite($fd, '<fonts x14ac:knownFonts="1" count="'.$this->fontsCount.'">');
		fwrite($fd, '	<font>');
		fwrite($fd, '		<sz val="'.$this->defaultFontSize.'"/>');
		fwrite($fd, '		<color theme="1"/>');
		fwrite($fd, '		<name val="'.$this->defaultFontName.'"/>');
		fwrite($fd, '		<family val="2"/>');
		if ($this->defaultFontName == 'MS Sans Serif') {
			fwrite($fd, '		<charset val="204"/>');
		} else if ($this->defaultFontName == 'Calibri') {
			fwrite($fd, '		<scheme val="minor"/>');
		} else {
			fwrite($fd, '		<charset val="204"/>');
		}
		fwrite($fd, '	</font>');
		if ($this->defaultStyle) {
			foreach ($this->defaultStyle as $style) {
				if (isset($style['sheet'])) {
					if (isset($style['font'])) {
						if (isset($style['font']['name']) && !empty($style['font']['name'])) $this->fontName = $style['font']['name'];
						if (isset($style['font']['size']) && !empty($style['font']['size'])) $this->fontSize = $style['font']['size'];
						if (isset($style['font']['color']) && !empty($style['font']['color'])) $this->fontColor = $style['font']['color'];
						if (isset($style['font']['bold']) && !empty($style['font']['bold'])) $this->fontStyles .= '<b/>';
						if (isset($style['font']['italic']) && !empty($style['font']['italic'])) $this->fontStyles .= '<i/>';
						if (isset($style['font']['underline']) && !empty($style['font']['underline'])) $this->fontStyles .= '<u/>';

						fwrite($fd, '	<font>');
						if ($this->fontStyles) fwrite($fd, '		'.$this->fontStyles);
						fwrite($fd, '		<sz val="'.$this->fontSize.'"/>');
						if ($this->fontColor) {
							fwrite($fd, '		<color rgb="FF'.$this->fontColor.'"/>');
						} else {
							fwrite($fd, '		<color theme="1"/>');
						}
						if ($this->fontName) fwrite($fd, '		<name val="'.$this->fontName.'"/>');
						fwrite($fd, '		<family val="2"/>');
						if ($this->fontName == 'MS Sans Serif') {
							fwrite($fd, '		<charset val="204"/>');
						} else if ($this->fontName == 'Calibri') {
							fwrite($fd, '		<scheme val="minor"/>');
						} else {
							fwrite($fd, '		<charset val="204"/>');
						}
						fwrite($fd, '	</font>');
					}
					$this->fontStyles = '';
				}
			}
		}
		fwrite($fd, '</fonts>');
		if ($this->defaultStyle) {
			foreach ($this->defaultStyle as $style) {
				if (isset($style['sheet'])) {
					if (isset($style['fill'])) $this->fillsCount++;
				}
			}
		}
		fwrite($fd, '<fills count="'.$this->fillsCount.'">');
		fwrite($fd, '	<fill><patternFill patternType="none"/></fill>');
		fwrite($fd, '	<fill><patternFill patternType="gray125"/></fill>');
		if ($this->defaultStyle) {
			foreach ($this->defaultStyle as $style) {
				if (isset($style['sheet'])) {
					if (isset($style['fill'])) {
						if (isset($style['fill']['color'])) $this->fillColor = $style['fill']['color'];
						fwrite($fd, '	<fill>');
						fwrite($fd, '		<patternFill patternType="solid">');
						fwrite($fd, '			<fgColor rgb="FF'.$this->fillColor.'"/>');
						fwrite($fd, '			<bgColor indexed="64"/>');
						fwrite($fd, '		</patternFill>');
						fwrite($fd, '	</fill>');
					}
				}
			}
		}
		fwrite($fd, '</fills>');
		if ($this->defaultStyle) {
			foreach ($this->defaultStyle as $style) {
				if (isset($style['sheet'])) {
					if (isset($style['border'])) $this->bordersCount++;
				}
			}
		}
		fwrite($fd, '<borders count="'.$this->bordersCount.'">');
		fwrite($fd, '	<border>');
		fwrite($fd, '		<left/><right/><top/><bottom/><diagonal/>');
		fwrite($fd, '	</border>');
		if ($this->defaultStyle) {
			foreach ($this->defaultStyle as $style) {
				if (isset($style['sheet'])) {
					if (isset($style['border'])) {
						if (isset($style['border']['style'])) $this->bordersStyle = ' style="'.$style['border']['style'].'"';
						if (isset($style['border']['color'])) $this->bordersColor = '<color rgb="FF'.$style['border']['color'].'"/>';
						fwrite($fd, '	<border>');
						fwrite($fd, '		<left'.$this->bordersStyle.'>'.$this->bordersColor.'</left>');
						fwrite($fd, '		<right'.$this->bordersStyle.'>'.$this->bordersColor.'</right>');
						fwrite($fd, '		<top'.$this->bordersStyle.'>'.$this->bordersColor.'</top>');
						fwrite($fd, '		<bottom'.$this->bordersStyle.'>'.$this->bordersColor.'</bottom>');
						fwrite($fd, '		<diagonal/>');
						fwrite($fd, '	</border>');
					}
				}
			}
		}
		fwrite($fd, '</borders>');
		fwrite($fd, 	'<cellStyleXfs count="1">');
		fwrite($fd,		'<xf borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		fwrite($fd, 	'</cellStyleXfs>');
		$this->stylesCount += count($this->defaultStyle);
		fwrite($fd, 	'<cellXfs count="'.$this->stylesCount.'">');
 		$this->defaultWrapText = ($this->defaultWrapText) ? '1' : '0';
		fwrite($fd, 		'<xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"><alignment wrapText="'.$this->defaultWrapText.'" vertical="'.$this->defaultVerticalAlign.'" horizontal="'.$this->defaultHorizontalAlign.'"/></xf>');
		if ($this->defaultStyle) {
			foreach ($this->defaultStyle as $style) {
				if (isset($style['sheet'])) {
					if (isset($style['font'])) {
						$font_Id = $this->fontId += 1;
					} else {
						$font_Id = 0;
					}
					if (isset($style['fill'])) {
						$fill_Id = $this->fillId += 1;
					} else {
						$fill_Id = 0;
					}
					if (isset($style['border'])) {
						$border_Id = $this->borderId += 1;
					} else {
						$border_Id = 0;
					}
					if (isset($style['wrapText'])) {
						$wrapText = ($style['wrapText']) ? '1' : '0';
					} else {
						$wrapText = $this->defaultWrapText;
					}

					$format_Id = (isset($style['format'])) ? $style['format'] : '0';

					if (isset($style['verticalAlign'])) {
						$verticalAlign = $style['verticalAlign'];
					} else {
						$verticalAlign = $this->defaultVerticalAlign;
					}
					if (isset($style['horizontalAlign'])) {
						$horizontalAlign = $style['horizontalAlign'];
					} else {
						$horizontalAlign = $this->defaultHorizontalAlign;
					}
					fwrite($fd, 		'<xf borderId="'.$border_Id.'" fillId="'.$fill_Id.'" fontId="'.$font_Id.'" numFmtId="'.$format_Id.'" xfId="0" applyFill="1">');
					fwrite($fd, 			'<alignment wrapText="'.$wrapText.'" vertical="'.$verticalAlign.'" horizontal="'.$horizontalAlign.'"/>');
					fwrite($fd, 		'</xf>');
				}
			}
		}
		fwrite($fd, 	'</cellXfs>');
		fwrite($fd, 	'<cellStyles count="1">');
		fwrite($fd, 		'<cellStyle xfId="0" builtinId="0" name="Normal"/>');
		fwrite($fd, 	'</cellStyles>');
		fwrite($fd, 	'<dxfs count="0"/>');
		fwrite($fd, 	'<tableStyles count="0" defaultPivotStyle="PivotStyleMedium9" defaultTableStyle="TableStyleMedium2"/>');
		fwrite($fd, 	'<extLst>');
		fwrite($fd, 	'<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}">');
		fwrite($fd, 		'<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>');
		fwrite($fd, 	'</ext>');
		fwrite($fd, 	'</extLst>');
		fwrite($fd, '</styleSheet>');
		fclose($fd);

		return $tempfile;
	}

	protected function setSharedString($v)
	{
		if (isset($this->shared_strings[$v]))
		{
			$string_value = $this->shared_strings[$v];
		}
		else
		{
			$string_value = count($this->shared_strings);
			$this->shared_strings[$v] = $string_value;
		}
		$this->shared_string_count++;//non-unique count
		return $string_value;
	}

	protected function writeSharedStringsXML()
	{
		$tempfile = $this->tempFilename();
		$fd = fopen($tempfile, "w+");
		if ($fd===false) { self::log("write failed in ".__CLASS__."::".__FUNCTION__."."); return; }
		
		fwrite($fd,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");
		fwrite($fd,'<sst count="'.($this->shared_string_count).'" uniqueCount="'.count($this->shared_strings).'" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
		foreach($this->shared_strings as $s=>$c)
		{
			fwrite($fd,'<si><t>'.self::xmlspecialchars($s).'</t></si>');
		}
		fwrite($fd, '</sst>');
		fclose($fd);
		return $tempfile;
	}

	protected function buildAppXML()
	{
		$app_xml="";
		$app_xml.='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
		$app_xml.='<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime></Properties>';
		return $app_xml;
	}

	protected function buildCoreXML()
	{
		$core_xml="";
		$core_xml.='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
		$core_xml.='<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
		$core_xml.='<dcterms:created xsi:type="dcterms:W3CDTF">'.date("Y-m-d\TH:i:s.00\Z").'</dcterms:created>';//$date_time = '2013-07-25T15:54:37.00Z';
		$core_xml.='<dc:creator>'.self::xmlspecialchars($this->author).'</dc:creator>';
		$core_xml.='<cp:revision>0</cp:revision>';
		$core_xml.='</cp:coreProperties>';
		return $core_xml;
	}

	protected function buildRelationshipsXML()
	{
		$rels_xml="";
		$rels_xml.='<?xml version="1.0" encoding="UTF-8"?>'."\n";
		$rels_xml.='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
		$rels_xml.='<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>';
		$rels_xml.='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
		$rels_xml.='<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>';
		$rels_xml.="\n";
		$rels_xml.='</Relationships>';
		return $rels_xml;
	}

	protected function buildWorkbookXML()
	{
		$workbook_xml="";
		$workbook_xml.='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
		$workbook_xml.='<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
		$workbook_xml.='<fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>';
		$workbook_xml.='<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>';
		$workbook_xml.='<sheets>';
		foreach($this->sheets_meta as $i=>$sheet_meta) {
			$workbook_xml.='<sheet name="'.self::xmlspecialchars($sheet_meta['sheetname']).'" sheetId="'.($i+1).'" state="visible" r:id="rId'.($i+2).'"/>';
		}
		$workbook_xml.='</sheets>';
		$workbook_xml.='<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>';
		return $workbook_xml;
	}

	protected function buildWorkbookRelsXML()
	{
		$wkbkrels_xml="";
		$wkbkrels_xml.='<?xml version="1.0" encoding="UTF-8"?>'."\n";
		$wkbkrels_xml.='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
		$wkbkrels_xml.='<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
		foreach($this->sheets_meta as $i=>$sheet_meta) {
			$wkbkrels_xml.='<Relationship Id="rId'.($i+2).'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/'.($sheet_meta['xmlname']).'"/>';
		}
		if (!empty($this->shared_strings)) {
			$wkbkrels_xml.='<Relationship Id="rId'.(count($this->sheets_meta)+2).'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>';
		}
		$wkbkrels_xml.="\n";
		$wkbkrels_xml.='</Relationships>';
		return $wkbkrels_xml;
	}

	protected function buildContentTypesXML()
	{
		$content_types_xml="";
		$content_types_xml.='<?xml version="1.0" encoding="UTF-8"?>'."\n";
		$content_types_xml.='<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
		$content_types_xml.='<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
		$content_types_xml.='<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
		foreach($this->sheets_meta as $i=>$sheet_meta) {
			$content_types_xml.='<Override PartName="/xl/worksheets/'.($sheet_meta['xmlname']).'" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
		}
		if (!empty($this->shared_strings)) {
			$content_types_xml.='<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
		}
		$content_types_xml.='<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
		$content_types_xml.='<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
		$content_types_xml.='<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
		$content_types_xml.='<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
		$content_types_xml.="\n";
		$content_types_xml.='</Types>';
		return $content_types_xml;
	}

	//------------------------------------------------------------------
	/*
	 * @param $row_number int, zero based
	 * @param $column_number int, zero based
	 * @return Cell label/coordinates, ex: A1, C3, AA42
	 * */
	public static function xlsCell($row_number, $column_number)
	{
		$n = $column_number;
		for($r = ""; $n >= 0; $n = intval($n / 26) - 1) {
			$r = chr($n%26 + 0x41) . $r;
		}
		return $r . ($row_number+1);
	}
	//------------------------------------------------------------------
	public static function log($string)
	{
		file_put_contents("php://stderr", date("Y-m-d H:i:s:").rtrim(is_array($string) ? json_encode($string) : $string)."\n");
	}
	//------------------------------------------------------------------
	public static function xmlspecialchars($val)
	{
		return str_replace("'", "&#39;", htmlspecialchars($val));
	}
	//------------------------------------------------------------------
	public static function array_first_key(array $arr)
	{
		reset($arr);
		$first_key = key($arr);
		return $first_key;
	}
	//------------------------------------------------------------------
	public static function convert_date_time($date_input) //thanks to Excel::Writer::XLSX::Worksheet.pm (perl)
	{
		$days    = 0;    # Number of days since epoch
		$seconds = 0;    # Time expressed as fraction of 24h hours in seconds
		$year=$month=$day=0;
		$hour=$min  =$sec=0;

		$date_time = $date_input;
		if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", $date_time, $matches))
		{
			list($junk,$year,$month,$day) = $matches;
		}
		if (preg_match("/(\d{2}):(\d{2}):(\d{2})/", $date_time, $matches))
		{
			list($junk,$hour,$min,$sec) = $matches;
			$seconds = ( $hour * 60 * 60 + $min * 60 + $sec ) / ( 24 * 60 * 60 );
		}

		//using 1900 as epoch, not 1904, ignoring 1904 special case
		
		# Special cases for Excel.
		if ("$year-$month-$day"=='1899-12-31')  return $seconds      ;    # Excel 1900 epoch
		if ("$year-$month-$day"=='1900-01-00')  return $seconds      ;    # Excel 1900 epoch
		if ("$year-$month-$day"=='1900-02-29')  return 60 + $seconds ;    # Excel false leapday

		# We calculate the date by calculating the number of days since the epoch
		# and adjust for the number of leap days. We calculate the number of leap
		# days by normalising the year in relation to the epoch. Thus the year 2000
		# becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
		$epoch  = 1900;
		$offset = 0;
		$norm   = 300;
		$range  = $year - $epoch;

		# Set month days and check for leap year.
		$leap = (($year % 400 == 0) || (($year % 4 == 0) && ($year % 100)) ) ? 1 : 0;
		$mdays = array( 31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 );

		# Some boundary checks
		if($year < $epoch || $year > 9999) return 0;
		if($month < 1     || $month > 12)  return 0;
		if($day < 1       || $day > $mdays[ $month - 1 ]) return 0;

		# Accumulate the number of days since the epoch.
		$days = $day;    # Add days for current month
		$days += array_sum( array_slice($mdays, 0, $month-1 ) );    # Add days for past months
		$days += $range * 365;                      # Add days for past years
		$days += intval( ( $range ) / 4 );             # Add leapdays
		$days -= intval( ( $range + $offset ) / 100 ); # Subtract 100 year leapdays
		$days += intval( ( $range + $offset + $norm ) / 400 );  # Add 400 year leapdays
		$days -= $leap;                                      # Already counted above

		# Adjust for Excel erroneously treating 1900 as a leap year.
		if ($days > 59) { $days++;}

		return $days + $seconds;
	}
	//------------------------------------------------------------------
}