# Powerpoint Sanitiser

Quickly and easily obfuscate the text in one or more Powerpoint files (PPTX only). Useful when sharing slide template layouts with others without sharing the content.

![Diagram](/screenshot.png)

# Installation

Ensure that .Net Core 3.1 is installed, then carry out the following steps:

```
git clone https://github.com/alirawashdeh/pptx-sanitiser.git
cd pptx-sanitiser
dotnet restore
```

# Usage

Place pptx files into the ```pptx-sanitiser``` folder or any subfolder. There is an example file in the ```pptx-sanitiser/pptx``` folder.
Run the application using the following:

```
dotnet run
```

A message will be outputted for each file processed:

```
Updated file: /Users/person/git/pptx-sanitiser/pptx/example.pptx 👌
```

## Warnings

- :warning: **Files are updated in situ** so make sure you copy the ```pptx``` files into the ```pptx-sanitiser folder``` rather than moving them

- :warning: Images are not sanitised during this process so make sure that you check each processed file for images (in the slides or master)

- :warning: Text within SmartArt