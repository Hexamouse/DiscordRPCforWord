# DiscordRPCforWord

**DiscordRPCforWord** is an experimental C++ utility that synchronizes your Microsoft Word activity with Discord Rich Presence in real-time.[web:2] It uses COM Late Binding for compatibility across Office versions like 2016, 2019, 2021, and 365.[web:9]

> **⚠️ EXPERIMENTAL!!** Use at your own risk—it interacts with Windows COM and active processes.[web:3]

## Key Features

- **Dynamic Document Detection**: Tracks the active Word document automatically.
- **English Status Display**: Shows professional status like "Page 3 of 15".
- **Late Binding COM**: Uses IDispatch for version-independent Word object model access.[web:6]
- **Performance Optimized**: Low CPU/memory use with C++17 optimizations.
- **x64 Target**: Built for modern Windows systems.

## Technical Specs

| Aspect          | Details                          |
|-----------------|----------------------------------|
| Language        | C++17                            |
| Core Tech       | Discord RPC SDK, Windows COM     |
| IDE             | Visual Studio 2022               |
| Platform        | x64                              |
| Runtime Library | Multi-threaded DLL (/MD)         |[web:4]

## Dependencies

- **Discord RPC SDK**: Place headers in `deps/include`, `.lib` in `deps/lib`.[web:5]
- **Windows SDK**: For COM and UTF-8 support (included in VS2022).

## Build Instructions

1. Clone the repo:  git clone https://github.com/username/DiscordRPCforWord.git
2. Open `DiscordRPCforWord.sln` in Visual Studio 2022.
3. Set **Release | x64** configuration.
4. In **Project Properties > C/C++ > Code Generation**: Set **Runtime Library** to **Multi-threaded DLL (/MD)**.
5. Build Solution (**Ctrl+Shift+B**).[web:4]

## Usage

1. Launch Microsoft Word and open a document.
2. Run `DiscordRPCforWord.exe`.
3. Your Discord status updates automatically with document name and page count.

**Note**: Ensure Discord is running with RPC enabled in your application settings.[web:11]

## License

MIT License. See the [LICENSE](LICENSE) file for details.

## Author

Developed by **Dzikri Maulana**. Feel free to contribute or report issues!
