# Import-Export-MAC-XMC-NBI

This program serves two purposes, to either import or export MAC addresses by End-System Group name into XMC using the NBI API call.

The program was created where a legacy network installation was being migrated to Extreme and adopting dynamic VLAN assignement through NAC.

All the MAC addresses and relative VLAN ID's were identified from the previous network. 

New End-System groups where determined, and MAC address assigned to their respective groups and entered into an Excel spreedsheet.

When first running the script it will ask you if you would like to import or export.

If exporting, the script will make an NBI call to XMC and pull all the MAC addresses, from all the groups, and export to an excel spreadshet that has the title row as the group name and the respective MAC addresses underneath. The exact same format is what is used for import. So first running and export will provide the intial template.

If importing, the script will list any Excel files that are in the same directory as the script, which you select, then asked which sheet to use. All the MAC addresses are extrapolated and compared to any that are in XMC already, and list of those that are not common is provided with the prompt y to create, or x to exit.

The next element will compare groups in the excel import with what is installed in XMC, and equally prompt to create or exit.
