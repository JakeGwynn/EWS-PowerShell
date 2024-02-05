# EWS-PowerShell

## Getting Started

These instructions will get you a copy of the project up and running on your local machine.

### Prerequisites

You need to have Git and PowerShell installed on your machine. Alternatively you can download the repository. However, it is recommended to use Git to clone the repo so that you can keep it up to date with Git. 

### Cloning the Repository

To clone this repository, open a terminal and navigate to the directory where you want the repository to be cloned. Then run the following command:

``` 
git clone https://github.com/JakeGwynn/EWS-PowerShell.git
```

### Importing the Module

After cloning the repository, you can import the module into your PowerShell session with the following command:
 
``` powershell
Import-Module "C:\path\to\EWS-PowerShell\EWS-PowerShell.psm1"
```

Use the example in ExampleCommands.ps1 if Git is installed in the default location.

## Using the Module

The module has a number of cmdlets (functions) that can be used to perform tasks in EWS. See ExampleCommands.ps1 for example usage. 


## License
Copyright 2024 Jackson Gwynn

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
