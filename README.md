
# EVPowerPointBundle
This is a Symfony2 Bundle helps you to read and write PowerPoint files, thanks to the PHPPowerPoint library

## Features
- Use easily the PHPPowerPoint library in Symfony 2

## Installation

In composer.json file, add :
```json
{
    "require": {
        "phpoffice/phppowerpoint": "dev-master",
        "ev/ev-powerpoint-bundle": "1.0.*@dev"
    }
}
```

In app/AppKernel.php file, add :
```php
public function registerBundles()
{
    return array(
        // ...
        new EV\PowerPointBundle\EVPowerPointBundle(),
        // ...
    );
}
```

## PHPPowerPoint's usage example

### Basic usage

In controller :
```php
// Acme\MainBundle\Controller\ExportController.php

public function powerpointAction() {
      
    $powerPointFactory = $this->get('ev_powerpoint');
        
    $objPowerPoint = $powerPointFactory->createPHPPowerPoint();

    // Create slide
    $currentSlide = $objPowerPoint->getActiveSlide();

    // Create a shape (text)
    $shape = $currentSlide->createRichTextShape()
          ->setHeight(300)
          ->setWidth(600)
          ->setOffsetX(170)
          ->setOffsetY(180);
    $textRun = $shape->createTextRun('Thank you for using EVPowerPointBundle!');
    $textRun->getFont()->setBold(true)
                       ->setSize(60)
                       ->setColor( new Color( 'FFE06B20' ) );

    $writer = $powerPointFactory->createWriter($objPowerPoint, 'PowerPoint2007');

    return $powerPointFactory->createStreamedResponseWithOptions($writer, array(
        'auto_headers' => true,
        'filename' => 'powerpoint_'.time()
    ));

}
```