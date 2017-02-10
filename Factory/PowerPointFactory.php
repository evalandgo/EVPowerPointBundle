<?php

namespace EV\PowerPointBundle\Factory;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Writer\Writerinterface;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007;
use PhpOffice\PhpPresentation\Writer\ODPresentation;


use Symfony\Component\HttpFoundation\StreamedResponse;


class PowerPointFactory {
    
    /**
     * 
     * @return \PhpOffice\PhpPresentation\PhpPowerpoint
     */
    public function createPHPPowerPoint() {
        return new PhpPresentation();
    }
    
    /**
     * 
     * @param type $name
     * @param type $properties
     */
    public function createObject($name, $properties = array()) {
        $class = 'PhpOffice\\PhpPresentation\\'.$name;
        
        $object = new $class();
        foreach($properties as $property => $value) {
            $method = 'set'.ucfirst($property);
            $object->$method($value);
        }
        
        return $object;
    }
    
    /**
     * 
     * @param type $name
     * @param type $properties
     * @return type
     */
    public function createShape($name, $properties = array()) {
        return $this->createObject('Sharpe\\'.$name, $properties);
    }
    
    /**
     * 
     * @param type $name
     * @param type $properties
     * @return type
     */
    public function createShared($name, $properties = array()) {
        return $this->createObject('Shared\\'.$name, $properties);
    }
    
    /**
     * 
     * @param type $name
     * @param type $properties
     * @return type
     */
    public function createStyle($name, $properties = array()) {
        return $this->createObject('Style\\'.$name, $properties);
    }
    
    /**
     * 
     * @param \PhpOffice\PhpPresentation\PhpPowerpoint $objPowerPoint
     * @param String $format
     * @return \PhpOffice\PhpPresentation\Writer\Writerinterface
     */
    public function createWriter(PhpPresentation $objPowerPoint, $format = 'PowerPoint2007') {
        return IOFactory::createWriter($objPowerPoint, $format);
    }
    
    /**
     * 
     * @param \PhpOffice\PhpPresentation\Writer\Writerinterface $writer
     * @return \Symfony\Component\HttpFoundation\StreamedResponse
     */
    public function createStreamedResponse(Writerinterface $writer) {
        return new StreamedResponse(function () use ($writer) {
            $writer->save('php://output');
        });
    }
    
    /**
     * 
     * @param \PhpOffice\PhpPresentation\Writer\Writerinterface $writer
     * @param Array $options
     * @return \Symfony\Component\HttpFoundation\StreamedResponse
     * @throws \Exception
     */
    public function createStreamedResponseWithOptions(Writerinterface $writer, $options = array()) {
        
        $response = $this->createStreamedResponse($writer);
        
        if ( isset($options['auto_headers']) && $options['auto_headers'] === true && isset($options['filename']) ) {
            
            if ( $writer instanceof ODPresentation ) {
                $response->headers->set('Content-Type', 'application/vnd.oasis.opendocument.presentation; charset=utf-8');
                $response->headers->set('Content-Disposition', 'attachment;filename='.$options['filename'].'.odp');
            }
            else if ( $writer instanceof PowerPoint2007 ) {
                //$response->headers->set('Content-Type', 'text/vnd.ms-powerpoint; charset=utf-8');
                $response->headers->set('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation; charset=utf-8');
                $response->headers->set('Content-Disposition', 'attachment;filename='.$options['filename'].'.pptx');
            }
            else {
                throw new \Exception("'auto_headers' option can't be used with this Writer. Only ODPPresentation and PowerPoint2007 are allowed");
            }
            
        }
        
        $response->headers->set('Pragma', 'public');
        $response->headers->set('Cache-Control', 'maxage=1');
        
        return $response;
    }
    
}

?>
