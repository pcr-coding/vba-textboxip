# VBA TextBoxIP
[![License: MIT](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://opensource.org/licenses/gpl-3.0)

Create a TextBox in a UserForm that takes IP addresses as input.

## Features

* Dots are inserted automatically each 3 digits. If user enters dots manually double dots are prevented.
* Editing between dots is possible.
* Validation for block size of 3 digits maximum.
* Validation for block value maxiumum 255.
* Accepts only numbers and dots as input.
* Set font automatically to "Consolas" for monospaced numbers.
* Supports copy/paste from clipboard. Invalid IPs get corrected.


## Examples

### Initialize 2 TextBoxes in a UserForm as IP boxes

```vba
Option Explicit

Private m_CollectionOfIPboxes As Collection

Private Sub UserForm_Initialize()
    Set m_CollectionOfIPboxes = New Collection
    
    m_CollectionOfIPboxes.Add New TextBoxIP, "TextBox1"
    Set m_CollectionOfIPboxes("TextBox1").TextBox = Me.TextBox1
    
    m_CollectionOfIPboxes.Add New TextBoxIP, "TextBox2"
    Set m_CollectionOfIPboxes("TextBox2").TextBox = Me.TextBox2
End Sub
```

## Unit Testing
Not included yet.
