<?php

namespace chaslain;

use Exception;
use SimpleXMLElement;
use ZipArchive;

/**
 * 
 * @package common\components\docx
 */
class DocXReplace
{

   /**
    * 
    * @var ZipArchive
    */
   private $zip;
   
   /**
    * 
    * @var SimpleXMLElement
    */
   private $documentXMLElement;

   public function __destruct()
   {
      if ($this->zip)
      {
         $this->zip->close();
      }
   }

   public function replace($searchFor, $replace)
   {
      $elementContents = $this->getElementContent($this->documentXMLElement);

      $longString = "";

      foreach ($elementContents as $i=>$oneElement)
      {
         $longString .= $oneElement["data"];
      }

      if (($pos = strpos($longString, $searchFor)) !== false)
      {
         $this->replaceInMap($pos, $elementContents, strlen($searchFor), $replace);
         $this->update($elementContents);
      }
   }

   public function save()
   {
      $this->zip->deleteName("word/document.xml");
      $this->zip->addFromString("word/document.xml", $this->documentXMLElement->asXML());
      // chmod($this->og_location, 777);
   }

   /**
    * 
    * @param mixed $position 
    * @param mixed $contents 
    * @param mixed $searchLength 
    * @param mixed $replacement 
    * @return void 
    */
   private function replaceInMap($position, &$contents, $searchLength, $replacement)
   {
      $runningPosition  = 0;
      $i                = 0;
      
      $lengthMap = [];


      // loop until we reach the index where the replacement begins
      while ($runningPosition < $position)
      {
         $runningPosition += strlen($contents[$i]["data"]);
         $lengthMap[$i] = strlen($contents[$i]["data"]);
         $i++;
      }

      // get the start positoin of this index
      $positionOfThisSegment = array_sum($lengthMap);
      
      $difference = 0;
      
      if ($runningPosition > $position)
      {
         $i--;
         $positionOfThisSegment = array_sum($lengthMap) - $lengthMap[$i];

         // the difference between the start of this index and the location of the replacement beginning
         $difference = $position - $positionOfThisSegment;
      }
      
      
      // now , just splice
      $beginningOfWord = $difference;
      $endOfWord = $beginningOfWord+$searchLength;
      
      $ogLength = strlen($contents[$i]["data"]);
      $start = substr($contents[$i]["data"], 0, $beginningOfWord);
      $end = substr($contents[$i]["data"], $endOfWord);
      $contents[$i]["data"] = $start . $replacement . $end;
      $i++;
      
      $endOfWord -= $ogLength;

      // until the end of the word falls below 0, remove all the trailing
      while ($endOfWord > 0)
      {
         $ogLength = strlen($contents[$i]["data"]);
         $end = substr($contents[$i]["data"], $endOfWord);
   
   
   
         $contents[$i]["data"] = $end;
         $i++;
         $endOfWord -= $ogLength; // subtract the end of the word so we don't remove too much
      }
   }

   private function update($contents)
   {
      $map = $this->formatForUpdate($contents);
     
      
      $this->updateElementContent($this->documentXMLElement, $map);

   }

   /**
    * 
    * @param SimpleXMLElement $element 
    * @param mixed $map 
    * @param int $i 
    */
   private function updateElementContent($element, $map)
   {
      foreach ($map as $key=>$value)
      {
         $key = explode("_", $key);
         
         $this->updateSingleElement($element, $key, $value);
      }
   }

   private function updateSingleElement($element, array $key, $value, $index = 0)
   {
      if ($index == count($key)-1)
      {
         $item = $key[(int)$index];
         $element[0] = $value;
         return;
      }


      if (is_numeric($key[$index]))
      {
         $item = $key[$index];
         $a = $element[(int)$item];
         
         $this->updateSingleElement($a, $key, $value, $index+1);
      }
      else
      {
         $val = $key[$index];

         $this->updateSingleElement($element->children("w", true)->$val, $key, $value, $index+1);
      }
   }

   private function formatForUpdate($contents)
   {
      $result = [];
      foreach ($contents as $one)
      {
         $result[implode("_", $one["loc"])] = $one["data"];
      }

      return $result;
   }


   /**
    * 
    * @param SimpleXMLElement $element 
    * @param mixed $search 
    * @param mixed $replace 
    * @return array 
    */
   private function getElementContent($element, $location = [])
   {
      $result = [];

      $elementMap = [];

      foreach ($element->children("w", true) as $i=>$oneElement)
      {
         $t = $location;
         $t[] = $i;
         $t[] = $this->incrementElementMap($elementMap, $i);


         if ($oneElement->count() > 0)
         {
            $result = array_merge($result, $this->getElementContent($oneElement, $t));
         }
         else
         {
            if (strlen($oneElement) > 0)
            {
               $result[] = ["data" => (string)$oneElement, "loc" => $t];
            }
         }
      }

      return $result;
   }

   private function incrementElementMap(&$map, $element)
   {
      if (!isset($map[$element]))
      {
         return 
         $map[$element] = 0;
      }

      $map[$element]++;

      return $map[$element];

   }

   private function getDocumentContent()
   {
      $file = $this->zip->getFromName("word/document.xml");
      $this->documentXMLElement = simplexml_load_string($file);

      if ($this->documentXMLElement === null)
      {
         throw new Exception("word document contains invalid xml???");
      }

   }


   public function open($docx)
   {
      $this->zip = new ZipArchive;
      $this->og_location = $docx;

      if (!file_exists($docx))
      {
         throw new Exception("{$docx} does not exist.");
      }

      $this->zip->open($docx);
      $this->getDocumentContent($docx);
   }
}