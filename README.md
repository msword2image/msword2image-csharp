# MSWord2Image-C&#35;

This library allows you to quickly convert Microsoft Word documents to image through [msword2image.com](http://msword2image.com) using C# for free!

## Demo

Example conversion: From [demo.docx](http://msword2image.com/docs/demo.docx) to [output.png](http://msword2image.com/docs/demoOutput.png). 

Note that you can try this out by visting [msword2image.com](http://msword2image.com) and clicking "Want to convert just one?"

## Installation

You can simply download [MsWordToImage.dll](https://github.com/msword2image/msword2image-csharp/raw/master/MsWordToImage/pack/MsWordToImage.dll) that includes all dependencies and referene this DLL.

## Usage

### 1. Convert from Word document URL to JPEG file

```csharp
MsWordToImageConvert convert = new MsWordToImageConvert(apiUser, apiKey);
convert.fromURL("http://msword2image.com/docs/demo.docx");
convert.toFile("output.jpeg");
// Please make sure output file is writable by your PHP process.
```

### 2. Convert from Word document URL to base 64 JPEG string

```csharp
```

### 3. Convert from Word file to JPEG file

```csharp
// Please make sure output file is writable and input file is readable by your PHP process.
```

### 4. Convert from Word file to base 64 encoded JPEG string

```csharp
// Please make sure input file is readable by your PHP process.
```

### 5. Convert from Word file to base 64 encoded GIF string

```csharp
// Please make sure input file is readable by your PHP process.
```

### 6. Convert from Word file to base 64 encoded PNG string

```csharp
// Please make sure input file is readable by your PHP process.
```

## Supported file formats

<table>
  <tbody>
    <tr>
      <td>Input\Output</td>
      <td>PNG</td>
      <td>GIF</td>
      <td>JPEG</td>
    </tr>
    <tr>
      <td>DOC</td>
      <td>✔</td>
      <td>✔</td>
      <td>✔</td>
    </tr>
    <tr>
      <td>DOCX</td>
      <td>✔</td>
      <td>✔</td>
      <td>✔</td>
    </tr>
  </tbody>
</table>
