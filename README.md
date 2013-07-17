# PHPWord - OpenXML - Read, Write and Create Word documents in PHP
PHPWord is a library written in PHP that create word documents.
No Windows operating system is needed for usage because the result are docx files (Office Open XML) that can be opened by all major office software.

## PHPWord 2.0
This is a fork of upstream project of [PHPWord](https://github.com/PHPOffice/PHPWord); henceforth referred to as *PHPWord 1.0*.

While this is distributed as *PHPWord 2.0* it strives to maintain backwards compatibility with *1.0* some changes were made to the API that will necessitate breakage.

### Why PHPWord 2.0?

* The upstream project does not seem to be maintained (outstanding pull-requests 8+months old).
    * [#13 - This is 2012, assume already have UTF8](https://github.com/PHPOffice/PHPWord/pull/13)
    * [#14 - Support for tab stops in Word 2007 documents](https://github.com/PHPOffice/PHPWord/pull/14)
    * [#15 - Support Multiple headers](https://github.com/PHPOffice/PHPWord/pull/15)

* The code does not adhere to any of the Frame Interop Group coding standards (or any standards from what I can tell).

    * [PSR-0](https://github.com/php-fig/fig-standards/blob/master/accepted/PSR-0.md)
    * [PSR-1](https://github.com/php-fig/fig-standards/blob/master/accepted/PSR-1-basic-coding-standard.md)
    * [PSR-2](https://github.com/php-fig/fig-standards/blob/master/accepted/PSR-2-coding-style-guide.md)

* Code is not DRY and has a lot of cruft (e.g., `diff --unified src/PHPWord/Writer/Word2007/Header.php src/PHPWord/Writer/Word2007/Footer.php`).

* There are no tests (yes, I know there are examples).

## Want to contribute?
Patches are always welcome. Fork and submit pull requests. I will help work with you to get the code merged into the codebase.

A couple of things to keep in mind before your patch will be accepted.

1. Everything in the `src` directory (excluding examples) **shall** adhere to the PSR guidelines (upto [PSR-2](https://github.com/php-fig/fig-standards/blob/master/accepted/PSR-2-coding-style-guide.md)).
2. If it is a new feature request, tests **shall** be implemented.
3. All tests **shall** pass before being merged.
4. Simple example scripts that illustrate how to utilize the code **should** be included.

## Checking PSR-2 Compliance

This is how I am going to check PSR-2 compliance. `phpcs --standard=PSR2 --ignore=src/Examples/ src/`

## Installing `phpcs`

`pear install PHP_CodeSniffer`

## License
PHPWord is licensed under [LGPL (GNU LESSER GENERAL PUBLIC LICENSE)](https://github.com/RLovelett/PHPWord/blob/master/license.md)