![Banner](images/banner.jpg)

# vbs_convert_LF_CRLF

> Convert from LF to CRLF and vice-versa

Description

## Table of Contents

- [Install](#install)
- [Usage](#usage)
- [See](#see)
- [License](#license)

## Install

Just grab a copy of this script and store in on your computer.

## Usage

Start a DOS prompt, go in the folder where you've save the script and run it like, f.i.,

```
cscript convert.vbs C:\Data TXT CRLF
```

or

```
cscript convert.vbs . TXT CRLF
```

The script ask three parameters:

1. The folder to process. Specify a folder or just type a dot (`.`) for the current folder,
2. The extension of files to process. If you wish to process every `.txt` files, just type `TXT` (case insensitive),
3. Then the desired line endings. If you wish that all files have Windows line endings, type `CRLF`. If a file is with Unix line endings, the file will be converted and rewrite. If you wish the opposite, type `LF` for the third parameter.

Note: the script contains a `SILENT_MODE` constant; the default value is `false`. If you wish that the script does his job silently, update the constant like this:

```vbnet
Const SILENT_MODE = false
```

## See

Part of this script is based on the work of [Stephen Millard](http://www.thoughtasylum.com/blog/2015/6/28/vbscript-flip-line-endings.html)

## License

[MIT](LICENSE)
