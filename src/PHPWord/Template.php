<?php

/**
 * PHPWord
 *
 * Copyright (c) 2011 PHPWord
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPWord
 * @package    PHPWord
 * @copyright  Copyright (c) 010 PHPWord
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    Beta 0.6.3, 08.07.2011
 */

/**
 * PHPWord_DocumentProperties
 *
 * @category   PHPWord
 * @package    PHPWord
 * @copyright  Copyright (c) 2009 - 2011 PHPWord (http://www.codeplex.com/PHPWord)
 */
class PHPWord_Template
{

    /**
     * ZipArchive
     *
     * @var ZipArchive
     */
    private $_objZip;

    /**
     * Temporary Filename
     *
     * @var string
     */
    private $_tempFileName;

    /**
     * Document XML
     *
     * @var string
     */
    private $_documentXML;

    /**
     * Tables from document XML
     * @var array
     */
    private $_tables;

    public function getDocumentXML()
    {
        return $this->_documentXML;
    }

    public function setDocumentXML($documentXML)
    {
        $this->_documentXML = $documentXML;
    }

    public function getTables()
    {
        return $this->_tables;
    }

    /**
     * Create a new Template Object
     *
     * @param string $strFilename
     */
    public function __construct($strFilename)
    {
        $path = dirname($strFilename);
        //$this->_tempFileName = $path . DIRECTORY_SEPARATOR . time() . '.docx';
        $this->_tempFileName = $path . DIRECTORY_SEPARATOR . microtime() . '.docx';

        copy($strFilename, $this->_tempFileName); // Copy the source File to the temp File

        $this->_objZip = new ZipArchive();
        $this->_objZip->open($this->_tempFileName);

        $this->_documentXML = $this->_objZip->getFromName('word/document.xml');
    }

    /**
     * Set Template values
     *
     * @param mixed $values
     */
    public function setValues($values)
    {
        Utils::convertArrayDataToISO($values);

        foreach ($values as $key => $value)
        {
            $this->setValue($key, $value);
        }
    }

    /**
     * Set a Template value
     *
     * @param mixed $search
     * @param mixed $replace
     * @param boolean $utf8_encode
     */
    public function setValue($search, $replace, $utf8_encode = FALSE) {
        $pattern = '|\$\{([^\}]+)\}|U';
        preg_match_all($pattern, $this->_documentXML, $matches);
        $openedTagPattern = '/<[^>]+>/';
        $closedTagPattern = '/<\/[^>]+>/';
        foreach ($matches[0] as $value) {
            $modificado = preg_replace($openedTagPattern, '', $value);
            $modificado = preg_replace($closedTagPattern, '', $modificado);
            $this->_documentXML = str_replace($value, $modificado, $this->_documentXML);
        }

        if (substr($search, 0, 1) !== '${' && substr($search, -1) !== '}') {
            $search = '${'.$search.'}';
        }

		preg_match_all('/\$[^\$]+?}/', $this->_documentXML, $matches);

		for ($i = 0; $i < count($matches[0]); $i++) {
			$matches_new[$i] = preg_replace('/(<[^<]+?>)/', '', $matches[0][$i]);
			$this->_documentXML = str_replace($matches[0][$i], $matches_new[$i], $this->_documentXML);
		}

        if(!is_array($replace) AND $utf8_encode !== FALSE) {
            $replace = utf8_encode($replace);
        }

        // Deletes xml-tags inside ${} - Source: http://phpword.codeplex.com/discussions/268012
        preg_match_all('/\$[^\$]+?}/', $this->_documentXML, $matches);

        for ($i = 0; $i < count($matches[0]); $i++) {
            $matches_new[$i] = preg_replace('/(<[^<]+?>)/', '', $matches[0][$i]);
            $this->_documentXML = str_replace($matches[0][$i], $matches_new[$i], $this->_documentXML);
        }

        $this->_documentXML = str_replace($search, $replace, $this->_documentXML);
    }

    /**
     * Retrieve all tabs from XML
     */
    public function getTablesFromXML()
    {
        $results = array();
        $offset = 0;
        while (($start = strpos($this->_documentXML, '<w:tbl>', $offset)) !== false)
        {
            $offset = $start + 7;

            $end = strpos($this->_documentXML, '</w:tbl>', $offset) + 8;

            $offset = $end + 8;

            $tab = substr($this->_documentXML, $start, ($end - $start));

            $results[0][] = $tab;
        }
        $this->_tables = $results;
    }

    /**
     * Insert a list of xml tables into document
     * @param array $xmlTables
     */
    public function insertXMLTables(array $xmlTables)
    {
        foreach ($xmlTables as $id => $xmlTable)
        {
            $this->insertXMLTable($id, $xmlTable);
        }
    }

    /**
     * Replace ${tag} by specified xml table
     * @param string $tag
     * @param string $xml
     */
    public function insertXMLTable($tag, $xml)
    {
        $this->_documentXML = preg_replace('/<w:p(\s+[^>]+)*>(?:(?!<\/w:p>).)*' . $tag . '}.*?<\/w:p>/', $xml, $this->_documentXML);
    }

    /**
     * Replace image by $path/$imageName
     * @param string $path
     * @param string $imageName
     */
    public function replaceImage($path, $imageName)
    {
        $this->_objZip->deleteName('word/media/' . $imageName);
        $this->_objZip->addFile($path, 'word/media/' . $imageName);
    }

    /**
     * Save Template
     * @param string $strFilename
     */
    public function save($strFilename)
    {
        if (file_exists($strFilename))
        {
            unlink($strFilename);
        }

        $this->_objZip->addFromString('word/document.xml', $this->_documentXML);

        // Close zip file
        if ($this->_objZip->close() === false)
        {
            throw new Exception('Could not close zip file.');
        }

        rename($this->_tempFileName, $strFilename);
    }

    /**
     * Clone a table row
     * @see http://jeroen.is/phpword-templates-with-repeating-rows/
     *
     * @param mixed $search
     * @param mixed $numberOfClones
     */
    public function cloneRow($search, $numberOfClones) {
        if(substr($search, 0, 2) !== '${' && substr($search, -1) !== '}') {
            $search = '${'.$search.'}';
        }

        $tagPos 	 = strpos($this->_documentXML, $search);
        $rowStartPos = strrpos($this->_documentXML, "<w:tr", ((strlen($this->_documentXML) - $tagPos) * -1));
        $rowEndPos   = strpos($this->_documentXML, "</w:tr>", $tagPos) + 7;

        $result = substr($this->_documentXML, 0, $rowStartPos);
        $xmlRow = substr($this->_documentXML, $rowStartPos, ($rowEndPos - $rowStartPos));
        for ($i = 1; $i <= $numberOfClones; $i++) {
            $result .= preg_replace('/\$\{(.*?)\}/','\${\\1#'.$i.'}', $xmlRow);
        }
        $result .= substr($this->_documentXML, $rowEndPos);

        $this->_documentXML = $result;
    }

    /**
     * Get document's content as string
     * @return content of document
     */
    public function get()
    {
        if (is_file($this->_tempFileName))
        {
            //add document
            $this->_objZip->addFromString('word/document.xml', $this->_documentXML);

            // Close zip file
            $this->_objZip->close();

            return file_get_contents($this->_tempFileName);
        }
    }

    /**
     * Destruct temp file if exists
     */
    function __destruct()
    {
        if (is_file($this->_tempFileName))
        {
            unlink($this->_tempFileName);
        }
    }

}
?>
