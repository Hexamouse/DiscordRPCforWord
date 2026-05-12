# DiscordRPCforWord (EXPERIMENTAL!!)

DiscordRPCforWord is an experimental C++ utility designed to synchronize Microsoft Word activity with Discord Rich Presence in real-time. This project utilizes a COM Late Binding approach to ensure broad compatibility across various Microsoft Office versions without relying on static Type Library imports.

Note: This project is currently labeled as EXPERIMENTAL!!. Use it at your own risk, as it interacts directly with the Windows COM layer and active processes.

Key Features
Dynamic Document Detection: Automatically identifies and tracks the currently active document in Microsoft Word.

English Status Localization: Displays activity using professional English terminology (e.g., Page 3 of 15).

Late Binding COM Implementation: Leverages the IDispatch interface to interact with Microsoft Word's object model. This technique eliminates the need for #import directives, ensuring stability across different Office versions including 2016, 2019, 2021, and Office 365.

Performance Optimized: Compiled with high-performance optimizations to maintain a minimal CPU and memory footprint during background operation.

Technical Specifications
Language: C++17

Core Technologies:

Discord Rich Presence SDK

Windows Component Object Model (COM)

Development Environment: Visual Studio 2022

Target Architecture: x64

Dependencies
The following external libraries are required for successful compilation:

Discord RPC SDK: Headers must be located in deps/include and the .lib file in deps/lib.

Windows SDK: Required for COM functionality and UTF-8 string manipulation.

Build and Installation Instructions
Clone the repository:

Bash
git clone https://github.com/username/DiscordRPCforWord.git
Open the solution file (.sln) in Visual Studio 2022.

Set the build configuration to Release and the platform to x64.

Ensure the Runtime Library in the project properties is set to Multi-threaded DLL (/MD) to align with the Discord RPC dependencies.

Execute Build Solution (Ctrl+Shift+B).

Usage
Launch Microsoft Word and open the document you wish to track.

Run the DiscordRPCforWord.exe executable.

The application will automatically interface with Discord to update your profile status with the current document name and page count.

License
This project is licensed under the Apache License 2.0. For more details, please refer to the LICENSE file.

Developed by Mysticalx / Red Vellano
