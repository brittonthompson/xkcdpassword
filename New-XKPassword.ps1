function Start-FastSearch {
    <#
    .SYNOPSIS
        Drastically improves finding items within an object over using the where-object filter.
    .DESCRIPTION    
        Uses the C# list search assembly to search a given object field for a specific value.
    .PARAMETER Object
        PowerShell object to search.
    .PARAMETER Field
        The field within the PowerShell object to search.
    .PARAMETER Value
        The value to search for within the field.
    .EXAMPLE
        Start-FastSearch -Object $Obj -Field "Name" -Value "Joe"
    #>
    
    [CmdletBinding()]
    param(
        [System.Object]$Object,
        [string]$Field,
        [string]$Value
    )

    $Source = @"
 
using System;
using System.Management.Automation;
using System.Collections.Generic;
    
namespace FastSearch
{
    
    public static class Search
    {
        public static List<Object> Find(PSObject[] collection, string column, string data)
        {
            List<Object> results = new List<Object>();
            foreach(PSObject item in collection)
            {
                if (item.Properties[column].Value.ToString() == data) { results.Add(item); }
            }
    
            return results;
        }
    }
}
"@

    if (-not ([System.Management.Automation.PSTypeName]'FastSearch.Search').Type) {
        Add-Type -ReferencedAssemblies $Assem -TypeDefinition $Source -Language CSharp
    }

    [FastSearch.Search]::Find($Object, $Field, $Value)
}


function New-XKPassword {
    <#
    .SYNOPSIS
        Creates a random password with the given characteristics. Inspired by the XKCD Comic and xkpasswd.net.
    .DESCRIPTION    
        This function uses a required dictionary file to create a random memorable password. The dictionary should use common words for performance.
        While testing and creating the function the most common 1000 English words were used from https://www.ef.edu/english-resources/english-vocabulary/top-1000-words.
        The CSV is two columns, Word and StringLength, where the StringLength column is the number of letters in the Word column.
    .PARAMETER DictionaryFile
        Full path to the csv dictionary.
    .PARAMETER MinWordLength
        The minimum length of any given word defaults to 4 and will never be shorter than this number.
    .PARAMETER MaxWordLength
        The maximumn length of any given word defaults to 8 and will never be longer than this number.
    .PARAMETER WordCount
    The maximum number of words to concatenate together in the final password
    .NOTES
        XKCD Comic 936: https://xkcd.com/936/
        XKPasswd:       https://xkpasswd.net/
    #>
    [cmdletBinding()]
    [OutputType([string])]
    param( 
        [ValidateRange(1, 9)]
        [int]$MinWordLength = 4,
        [ValidateRange(1, 9)]        
        [int]$MaxWordLength = 8,
        [ValidateRange(1, 24)]        
        [int]$WordCount = 3, 
        [string]$DictionaryFile ="words.csv"
    )
    
    if (Test-Path $DictionaryFile) {
        #Check to see if the fast search assembly is loaded
        #Fast search usesa C# type definition included in the common.ps1
        if (Get-Command "Start-FastSearch" -errorAction SilentlyContinue) { $FastSearch = $True }
    
        #Get the dictionary file so we can get the max and min length if different from given
        $Dictionary = Import-Csv -Path $DictionaryFile

        #Get an ordered list of all the lengths so we can make sure we don't try to pick a word of length that doesn't exist
        $WordLengths = $Dictionary.StringLength | Select-Object -Unique | Where-Object { $_ -ge $MinWordLength -and $_ -le $MaxWordLength }

        #Create a list of random word lengths
        $Lengths = @()
        for ($i = 1; $i -le $WordCount; $i++) { $Lengths += $(Get-Random -InputObject $WordLengths) }
            
        #Get a list of random words with the given lengths
        $RandomWords = @()
        foreach ($L in $Lengths) {
            if ($FastSearch) {
                $WordsOfLength = ($(Start-FastSearch -Object $Dictionary -Field "StringLength" -Value $L)).Word
            }
            else {
                $WordsOfLength = ($Dictionary | Where-Object { $_.StringLength -eq $L }).Word
            }
            $RandomWords += $(Get-Random -InputObject $WordsOfLength)
        }
    
        #Generate the formatted password with symbols and numbers
        $Symbols = @('!', '@', '#', '$', '%', '^', '&', '*', '(', ')', '-', '_', '=', '+', '~')

        #Get a random symbol and double it for the first and last characters in the password
        $Outside = Get-Random -InputObject $Symbols
        $Outside = $Outside + $Outside

        #Get a random symbol to use as the word separator
        $Inside = Get-Random -InputObject $Symbols

        #Create the left side string of the password with the outside characters a random double digit number and the first inside character
        $Left = $Outside + $("{0:D2}" -f [int]$(Get-Random -Minimum 0 -Maximum 99)) + $Inside

        #Create the right side of the string with another random double digit number and the outside characters
        $Right = $("{0:D2}" -f [int]$(Get-Random -Minimum 0 -Maximum 99)) + $Outside

        #Concatenate our random passwords with the inside character and every second word converted to upper case
        $Middle = ""
        for ($i = 0; $i -lt $RandomWords.Count; $i++) { 
            if ($i % 2 -eq 0) { $Word = $RandomWords[$i] } else { $Word = $RandomWords[$i].ToUpper() }
            $Middle += $Word, $Inside -join ""
        }

        #Concatenate and return the final password 
        return $($Left, $Middle, $Right -join "")
    }
    else {
        throw "Your dictionary file path is not valid"
    }
}

New-XKPassword