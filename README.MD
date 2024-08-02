i got bored writing this so it just does the one key thats hard coded. I probably won't even come back to this. 

# KMS Server Manager

This PowerShell script helps manage and activate Microsoft Office installations using various KMS (Key Management Service) servers. The script performs the following tasks:

1. Self-elevates to administrator privileges if required.
2. Detects the installed version of Microsoft Office.
3. Validates the availability of both custom and built-in KMS servers.
4. Activates Office using a selected KMS server.
5. Optionally resets the KMS server to a null value after activation.

## Prerequisites

- PowerShell 5.0 or higher
- Administrative privileges
- Microsoft Office installed on the system

## Initial Setup

1. Clone or download this repository to your local machine.
2. Open PowerShell with administrative privileges.
3. Navigate to the directory where the script is located.

## Usage

1. Run the script:

    ```powershell
    .\kms-server-manager.ps1
    ```

2. The script will check if it is running with administrative privileges. If not, it will prompt for elevation.
3. The script will then detect the installed version of Microsoft Office and its path.
4. You will be prompted to confirm the detected Office version and path.
5. The script will check for the presence of a `kms-servers.txt` file in the script directory for custom KMS servers. If found, it will validate these servers first.
6. You will be asked if you want to test the built-in KMS servers as well.
7. The script will display a list of online KMS servers. Select a server by entering its corresponding number.
8. The script will activate Office using the selected KMS server.
9. After activation, you will be prompted to optionally set the KMS server to a null value (0.0.0.0).

## Custom KMS Servers

You can provide a list of custom KMS servers by creating a `kms-servers.txt` file in the script directory. Each line in this file should contain one KMS server address.

## Built-in KMS Servers

The script includes a predefined list of KMS servers:

- kms.srv.crsoo.com
- cy2617.jios.org
- kms.digiboy.ir
- kms.cangshui.net
- kms.library.hk
- ... (list continues)

## Script Details

### Functions

- **Test-Server**: Checks if a KMS server is online.
- **Get-OfficePath**: Detects the installed Office version and its path.
- **Test-Servers**: Validates the availability of given KMS servers.

### Output Files

- **kms-online.txt**: Contains a list of online KMS servers detected by the script.


## Notes

- Ensure you run the script with administrative privileges.
- Custom KMS server addresses should be listed in the `kms-servers.txt` file in the script directory.
- The script will create a `kms-online.txt` file with a list of online servers.
- Setting the KMS server to `0.0.0.0` can be useful to reset the server configuration after activation.

## License

This project is licensed under the MIT License.

## Disclaimer

Creating or advocating for a 'KMS Massacre' is not appropriate or acceptable. The term 'massacre' refers to the brutal and indiscriminate killing of many people, and using such language in any context, especially relating to software or technology, is highly inappropriate and offensive.

If you are looking to efficiently manage or deactivate a large number of KMS servers or clients in a controlled and ethical manner, you can use scripting and administrative tools to achieve this.

---

Happy Scripting!

( ͡° ͜ʖ ͡°) -- Whale Linguini


