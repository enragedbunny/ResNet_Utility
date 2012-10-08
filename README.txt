***********************************************************************
Program: ResNet Utility
Created by: Johnny Keeton
Other Contributors: Alex Burgy-Vanhoose, Josh Back, CJ Highley
Last Edited: 2012-10-07
Written in: Autoit
Compatability: Windows Vista/7
***********************************************************************

                               About
***********************************************************************

ResNet Utility (RU) is a small program written for a university
computer support department. It's main function is to provide quick 
access to commonly used system information and provides access to 
various OS tools and utilities to reduce repetitive keystrokes and/or 
mouse movements.

The program is divided into tabs that correspond to different sets of 
features. The main tab is mostly to provide as much information about 
the current system as possible, while providing other resources 
(buttons) to provide quick access to frequently used programs or 
utilities.

                             Features
***********************************************************************

The ResNet Utility provides the following features

 * Collects system information
    * Operating System
    * System Architecture (32 or 64 bit)
    * Computer Brand (Dell, HP, etc.)
    * Computer Model
    * Total HDD Space and Free Space
    * RAM
    * Current Service Pack
    * Ethernet information (Description, MAC Address)
    * Wireless information (Description, MAC Address)
    * Serial Number
 * Provides several automated system repair functions
    * Resets network settings (Interfaces, winsock, etc)
    * Resets/Repairs firewall settings
    * SMART Test (for HDD testing)
    * Repair Windows Update
    * Repair Windows Permissions
    * Fix File Associations (beta)
 * Automates calling several windows and programs
    * Programs and Features
    * System Properties
    * ipconfig /all (through cmd.exe)

                            Change Log
***********************************************************************

Version 0.3.0 (2012.10.4)

* First version with README.txt
* Cleaned up code
* Commented code for readability
* Changed Version to 0.3.0
* Separated functions from main program