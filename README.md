# INI File Utility

A VB.NET class for managing INI files, preserving the original order of sections and keys, with methods for loading, saving, adding, removing, and renaming entries.

## Features

- **Load INI Files**: Read and parse INI files into structured objects.
- **Save INI Files**: Write INI file data back to disk, preserving the original order.
- **Manage Sections and Keys**: Add, remove, and rename sections and keys.
- **Preserve Order**: Maintain the original order of sections and keys, appending new entries to the end.

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/inifile-utility.git
    ```
2. Open the project in Visual Studio.
3. Build the solution to restore dependencies.

## Usage

### Loading an INI File

```vb.net
Dim ini As New DracLabs.IniFile()
ini.Load("path/to/yourfile.ini")
```

### Saving an INI File

```vb.net
ini.Save("path/to/yourfile.ini")
```

### Adding a Section

```vb.net
Dim section As DracLabs.IniFile.IniSection = ini.AddSection("NewSection")
```

### Adding a Key

```vb.net
Dim key As DracLabs.IniFile.IniSection.IniKey = section.AddKey("NewKey")
key.Value = "NewValue"
```

### Removing a Section

```vb.net
ini.RemoveSection("SectionName")
```

### Renaming a Section

```vb.net
ini.RenameSection("OldSectionName", "NewSectionName")
```

### Renaming a Key 

```vb.net
section.RenameKey("OldKeyName", "NewKeyName")
```

### Getting a Key Value

```vb.net
Dim value As String = ini.GetKeyValue("SectionName", "KeyName")
```

### Setting a Key Value

```vb.net
Dim success As Boolean = ini.SetKeyValue("SectionName", "KeyName", "NewValue")
```

### Removing All Sections

```vb.net
Dim success As Boolean = ini.RemoveAllSections()
```

## History
This code was originally created by Draco and Adultery and was shipped with the PinballX plugin. 5Cutters and Mike Da Spike modified it for use with the Database Manager and are now sharing it.

## Usage
Contributions are welcome! Please fork the repository and submit a pull request.

## License
This project is licensed under the MIT License. See the LICENSE file for details.
