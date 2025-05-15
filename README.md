# WiFiKeyViewer

**WiFiKeyViewer** is a lightweight and simple Python tool that allows users to retrieve saved Wi-Fi profiles and their associated passwords on Windows systems. This can be especially useful for network administrators, tech support, or users who have forgotten the passwords to networks they have previously connected to.

## Features

* Lists all saved Wi-Fi profiles on a Windows machine.
* Retrieves passwords for each saved network (when available).
* Outputs results directly to the console.
* Portable and does not require additional dependencies.

## Requirements

* Windows OS
* Python 3.x

## Usage

1. Clone the repository:

   ```bash
   git clone https://github.com/REZACODINNGCRAFTER/WiFiKeyViewer.git
   cd WiFiKeyViewer
   ```

2. Run the script:

   ```bash
   python WiFiKeyViewer.py
   ```

3. The script will display the list of saved Wi-Fi profiles and attempt to extract and display the passwords associated with them.

> ⚠️ **Note**: This script must be run with sufficient privileges to access the system's network configuration data.

## Disclaimer

This tool is intended for educational and personal use only. Unauthorized access to networks is strictly prohibited and may be illegal.

## Contributing

We welcome contributions from the community! If you're a Python developer and interested in enhancing WiFiKeyViewer, here are some ideas for contribution:

* Cross-platform support (e.g., Linux, macOS).
* Export functionality (e.g., save results to a file).
* GUI interface for ease of use.
* Error handling and logging improvements.

To contribute:

1. Fork the repository.
2. Create a new branch: `git checkout -b feature-name`
3. Make your changes.
4. Commit and push: `git commit -m "Add feature X" && git push origin feature-name`
5. Submit a pull request.

## License

This project is open-source and available under the MIT License.

---

Maintained by [REZACODINNGCRAFTER](https://github.com/REZACODINNGCRAFTER). Feel free to reach out with ideas, issues, or improvements!
