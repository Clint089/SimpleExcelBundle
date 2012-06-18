SimpleExcel
===========
This Bundle permits to read easly Excel-Files.


## INSTALLATION

1 Add the following entry to ``deps`` the run ``php bin/vendors install``.

``` yaml 
[OnemediaSimpleExcelBundle]
    git=https://github.com/Clint089/SimpleExcelBundle.git
    target=/bundles/Onemedia/SimpleExcelBundle
```

2 Register the bundle in ``app/AppKernel.php``

``` php
    $bundles = array(
        // ...
        new Onemedia\SimpleExcelBundle\OnemediaSimpleExcelBundle(),
    );
```

3  Register namespace in ``app/autoload.php``

``` php
    $loader->registerNamespaces(array(
         // ...
         'Onemedia'              => __DIR__.'/../vendor/bundles',
     ));
```



