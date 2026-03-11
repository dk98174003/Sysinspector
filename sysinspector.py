import os
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime


class SysInspectorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("SysInspector - Professional System Information Tool")

        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(base_dir, "sysinspector.ico")
            self.root.iconbitmap(icon_path)
        except Exception:
            pass

        self.root.geometry("1280x820")
        self.root.minsize(1024, 700)

        self.current_title = "Summary"
        self.current_command = None

        self.style = ttk.Style()
        try:
            self.style.theme_use("vista")
        except tk.TclError:
            try:
                self.style.theme_use("clam")
            except tk.TclError:
                pass

        self._build_ui()
        self.show_summary()

    def _build_ui(self) -> None:
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        main = ttk.Frame(self.root, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        main.rowconfigure(2, weight=1)
        main.columnconfigure(0, weight=0)
        main.columnconfigure(1, weight=1)

        self._build_header(main)
        self._build_toolbar(main)
        self._build_navigation(main)
        self._build_content(main)
        self._build_statusbar(main)

    def _build_header(self, parent: ttk.Frame) -> None:
        header = ttk.Frame(parent)
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        header.columnconfigure(1, weight=1)

        title = ttk.Label(
            header,
            text="SysInspector",
            font=("Segoe UI", 18, "bold")
        )
        title.grid(row=0, column=0, sticky="w")

        subtitle = ttk.Label(
            header,
            text="Professional Windows system information and inventory utility",
            font=("Segoe UI", 10)
        )
        subtitle.grid(row=0, column=1, sticky="w", padx=(12, 0), pady=(4, 0))

    def _build_toolbar(self, parent: ttk.Frame) -> None:
        toolbar = ttk.LabelFrame(parent, text="Toolbar", padding=8)
        toolbar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        toolbar.columnconfigure(10, weight=1)

        ttk.Button(toolbar, text="Summary", command=self.show_summary, width=14).grid(row=0, column=0, padx=4, pady=2)
        ttk.Button(toolbar, text="Refresh", command=self.refresh_current, width=14).grid(row=0, column=1, padx=4, pady=2)
        ttk.Button(toolbar, text="Copy Output", command=self.copy_output, width=14).grid(row=0, column=2, padx=4, pady=2)
        ttk.Button(toolbar, text="Save Output", command=self.save_output, width=14).grid(row=0, column=3, padx=4, pady=2)
        ttk.Button(toolbar, text="Full Report", command=self.show_full_report, width=14).grid(row=0, column=4, padx=4, pady=2)

        ttk.Label(toolbar, text="Filter:").grid(row=0, column=5, padx=(18, 4), pady=2, sticky="e")

        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", self.apply_filter)
        self.filter_entry = ttk.Entry(toolbar, textvariable=self.filter_var, width=28)
        self.filter_entry.grid(row=0, column=6, padx=4, pady=2, sticky="w")

        ttk.Button(toolbar, text="Clear Filter", command=self.clear_filter, width=14).grid(row=0, column=7, padx=4, pady=2)

    def _build_navigation(self, parent: ttk.Frame) -> None:
        nav_frame = ttk.LabelFrame(parent, text="Navigation", padding=8)
        nav_frame.grid(row=2, column=0, sticky="nsw", padx=(0, 10))
        nav_frame.rowconfigure(0, weight=1)
        nav_frame.columnconfigure(0, weight=1)

        self.nav_tree = ttk.Treeview(nav_frame, show="tree", height=28)
        self.nav_tree.grid(row=0, column=0, sticky="nsw")

        yscroll = ttk.Scrollbar(nav_frame, orient="vertical", command=self.nav_tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.nav_tree.configure(yscrollcommand=yscroll.set)

        system = self.nav_tree.insert("", "end", text="System", open=True)
        hardware = self.nav_tree.insert("", "end", text="Hardware", open=True)
        network = self.nav_tree.insert("", "end", text="Network", open=True)
        windows = self.nav_tree.insert("", "end", text="Windows", open=True)
        reports = self.nav_tree.insert("", "end", text="Reports", open=True)

        self.nav_actions = {}

        self._add_nav_item(system, "Summary", self.show_summary)
        self._add_nav_item(system, "BIOS Version", self.show_bios_version)
        self._add_nav_item(system, "Computer System", self.show_computersystem)
        self._add_nav_item(system, "CS Product", self.show_csproduct)
        self._add_nav_item(system, "Operating System", self.show_os)

        self._add_nav_item(hardware, "CPU", self.show_cpu)
        self._add_nav_item(hardware, "Desktop Monitor", self.show_desktopmonitor)
        self._add_nav_item(hardware, "Disk Drive", self.show_diskdrive)
        self._add_nav_item(hardware, "Local Disk", self.show_localdisk)
        self._add_nav_item(hardware, "Onboard Device", self.show_onboarddevice)

        self._add_nav_item(network, "Net Use", self.show_netuse)
        self._add_nav_item(network, "NIC", self.show_nic)
        self._add_nav_item(network, "NIC Config", self.show_nicconfig)
        self._add_nav_item(network, "NT Domain", self.show_ntdomain)

        self._add_nav_item(windows, "Printer", self.show_printer)
        self._add_nav_item(windows, "Process", self.show_process)
        self._add_nav_item(windows, "Product", self.show_product)
        self._add_nav_item(windows, "Service", self.show_service)
        self._add_nav_item(windows, "Share", self.show_share)
        self._add_nav_item(windows, "Startup", self.show_startup)
        self._add_nav_item(windows, "SysAccount", self.show_sysaccount)

        self._add_nav_item(reports, "Full Report", self.show_full_report)

        self.nav_tree.bind("<<TreeviewSelect>>", self.on_nav_select)

    def _add_nav_item(self, parent: str, text: str, callback) -> None:
        item_id = self.nav_tree.insert(parent, "end", text=text)
        self.nav_actions[item_id] = callback

    def _build_content(self, parent: ttk.Frame) -> None:
        content = ttk.Frame(parent)
        content.grid(row=2, column=1, sticky="nsew")
        content.rowconfigure(1, weight=1)
        content.columnconfigure(0, weight=1)

        self.content_header_var = tk.StringVar(value="Summary")
        self.content_subheader_var = tk.StringVar(value="Startup system overview")

        header_frame = ttk.Frame(content)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header_frame.columnconfigure(0, weight=1)

        ttk.Label(
            header_frame,
            textvariable=self.content_header_var,
            font=("Segoe UI", 15, "bold")
        ).grid(row=0, column=0, sticky="w")

        ttk.Label(
            header_frame,
            textvariable=self.content_subheader_var,
            font=("Segoe UI", 9)
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        output_frame = ttk.LabelFrame(content, text="Output", padding=8)
        output_frame.grid(row=1, column=0, sticky="nsew")
        output_frame.rowconfigure(0, weight=1)
        output_frame.columnconfigure(0, weight=1)

        self.output = tk.Text(
            output_frame,
            wrap="none",
            font=("Consolas", 10),
            undo=False,
            borderwidth=0
        )
        self.output.grid(row=0, column=0, sticky="nsew")

        yscroll = ttk.Scrollbar(output_frame, orient="vertical", command=self.output.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.output.configure(yscrollcommand=yscroll.set)

        xscroll = ttk.Scrollbar(output_frame, orient="horizontal", command=self.output.xview)
        xscroll.grid(row=1, column=0, sticky="ew")
        self.output.configure(xscrollcommand=xscroll.set)

    def _build_statusbar(self, parent: ttk.Frame) -> None:
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        status_frame.columnconfigure(0, weight=1)

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(status_frame, textvariable=self.status_var, anchor="w").grid(row=0, column=0, sticky="ew")

    def set_status(self, text: str) -> None:
        self.status_var.set(text)
        self.root.update_idletasks()

    def set_content_header(self, title: str, subtitle: str) -> None:
        self.content_header_var.set(title)
        self.content_subheader_var.set(subtitle)

    def clear_output(self) -> None:
        self.output.delete("1.0", "end")

    def append_output(self, title: str, content: str, subtitle: str = "") -> None:
        self.clear_output()
        self.set_content_header(title, subtitle or "System information view")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        self.output.insert("end", f"{'=' * 100}\n")
        self.output.insert("end", f"{title}\n")
        self.output.insert("end", f"Generated: {timestamp}\n")
        self.output.insert("end", f"{'-' * 100}\n")
        self.output.insert("end", f"{content.strip()}\n\n")
        self.output.see("1.0")

    def run_command(self, command: str) -> str:
        try:
            result = subprocess.run(
                command,
                shell=True,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace"
            )
            stdout = (result.stdout or "").strip()
            stderr = (result.stderr or "").strip()

            if result.returncode == 0 and stdout:
                return stdout
            if stderr:
                return f"Error:\n{stderr}"
            return "No output returned."
        except Exception as exc:
            return f"Exception:\n{exc}"

    def run_wmic(self, wmic_args: str) -> str:
        return self.run_command(f"wmic {wmic_args}")

    def run_simple_query(self, title: str, wmic_args: str, subtitle: str) -> None:
        self.current_title = title
        self.current_command = wmic_args
        self.set_status(f"Loading {title}...")
        result = self.run_wmic(wmic_args)
        self.append_output(title, result, subtitle)
        self.set_status(f"{title} loaded")

    def on_nav_select(self, event=None) -> None:
        selection = self.nav_tree.selection()
        if not selection:
            return
        item_id = selection[0]
        callback = self.nav_actions.get(item_id)
        if callback:
            callback()

    def refresh_current(self) -> None:
        if self.current_title == "Summary":
            self.show_summary()
            return

        if self.current_title == "Full Report":
            self.show_full_report()
            return

        if self.current_command:
            self.set_status(f"Refreshing {self.current_title}...")
            result = self.run_wmic(self.current_command)
            self.append_output(self.current_title, result, "Refreshed current view")
            self.set_status(f"{self.current_title} refreshed")
        else:
            self.show_summary()

    def copy_output(self) -> None:
        content = self.output.get("1.0", "end").strip()
        if not content:
            messagebox.showinfo("Nothing to copy", "There is no output to copy.")
            return

        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.set_status("Output copied to clipboard")
        messagebox.showinfo("Copied", "Output copied to clipboard.")

    def save_output(self) -> None:
        content = self.output.get("1.0", "end").strip()
        if not content:
            messagebox.showinfo("Nothing to save", "There is no output to save.")
            return

        filename = filedialog.asksaveasfilename(
            title="Save output",
            defaultextension=".txt",
            filetypes=[
                ("Text files", "*.txt"),
                ("All files", "*.*"),
            ],
            initialfile="sysinspector_output.txt"
        )
        if not filename:
            return

        try:
            with open(filename, "w", encoding="utf-8") as file:
                file.write(content)
            self.set_status(f"Saved: {filename}")
            messagebox.showinfo("Saved", f"Output saved successfully:\n{filename}")
        except Exception as exc:
            messagebox.showerror("Save error", f"Could not save file:\n{exc}")
            self.set_status("Save failed")

    def clear_filter(self) -> None:
        self.filter_var.set("")

    def apply_filter(self, *args) -> None:
        filter_text = self.filter_var.get().strip().lower()
        self.output.tag_remove("filter_match", "1.0", "end")

        if not filter_text:
            return

        start = "1.0"
        while True:
            pos = self.output.search(filter_text, start, stopindex="end", nocase=True)
            if not pos:
                break
            end = f"{pos}+{len(filter_text)}c"
            self.output.tag_add("filter_match", pos, end)
            start = end

        self.output.tag_config("filter_match", background="yellow", foreground="black")

    def _safe_value(self, wmic_args: str, default: str = "Unknown") -> str:
        result = self.run_wmic(wmic_args)
        lines = [line.strip() for line in result.splitlines() if line.strip()]
        if len(lines) >= 2:
            return lines[-1]
        if lines:
            return lines[-1]
        return default

    def show_summary(self) -> None:
        self.current_title = "Summary"
        self.current_command = None
        self.set_status("Building summary dashboard...")

        hostname = self._safe_value("computersystem get name")
        username = self._safe_value("computersystem get username")
        manufacturer = self._safe_value("computersystem get manufacturer")
        model = self._safe_value("computersystem get model")
        domain = self._safe_value("computersystem get domain")
        bios = self._safe_value("bios get BIOSversion")
        serial = self._safe_value("csproduct get identifyingnumber")
        uuid = self._safe_value("csproduct get uuid")
        os_name = self._safe_value("os get caption")
        os_version = self._safe_value("os get version")
        cpu = self._safe_value("cpu get name")
        cores = self._safe_value("cpu get numberofcores")
        monitor_width = self._safe_value("desktopmonitor get screenwidth")
        monitor_height = self._safe_value("desktopmonitor get screenheight")

        summary = f"""
SYSTEM SUMMARY

Computer Name       : {hostname}
User Name           : {username}
Manufacturer        : {manufacturer}
Model               : {model}
Domain              : {domain}

BIOS Version        : {bios}
Serial Number       : {serial}
UUID                : {uuid}

Operating System    : {os_name}
OS Version          : {os_version}

CPU                 : {cpu}
CPU Cores           : {cores}

Screen Resolution   : {monitor_width} x {monitor_height}

NOTES
- This dashboard provides a quick overview of the most important system details.
- Use the navigation panel on the left for detailed category views.
- Use Full Report for a complete inventory-style output.
""".strip()

        self.append_output("Summary", summary, "Startup system overview")
        self.set_status("Summary loaded")

    def show_bios_version(self) -> None:
        self.run_simple_query("BIOS Version", "bios get BIOSversion", "Firmware and BIOS version details")

    def show_computersystem(self) -> None:
        self.run_simple_query(
            "Computer System",
            "computersystem get name,username,manufacturer,model,domain",
            "Core computer identity and logged-on user information"
        )

    def show_cpu(self) -> None:
        self.run_simple_query(
            "CPU Information",
            "cpu get name,numberofcores,systemname",
            "Processor name and core count"
        )

    def show_csproduct(self) -> None:
        self.run_simple_query(
            "CS Product",
            "csproduct get identifyingnumber,name,uuid",
            "System product identification and UUID"
        )

    def show_desktopmonitor(self) -> None:
        self.run_simple_query(
            "Desktop Monitor",
            "desktopmonitor get screenheight,screenwidth,systemname",
            "Detected monitor resolution information"
        )

    def show_diskdrive(self) -> None:
        self.run_simple_query(
            "Disk Drive",
            "diskdrive get model,name,size",
            "Physical disk drive inventory"
        )

    def show_localdisk(self) -> None:
        self.run_simple_query(
            "Local Disk",
            "logicaldisk get name,providername,systemname,description",
            "Logical disk and mapping information"
        )

    def show_netuse(self) -> None:
        self.run_simple_query(
            "Net Use",
            "netuse get localname,username,remotepath",
            "Mapped network connection details"
        )

    def show_nic(self) -> None:
        self.run_simple_query(
            "NIC Information",
            "nic get adaptertype,macaddress,maxspeed,name,systemname,caption",
            "Network adapter hardware details"
        )

    def show_nicconfig(self) -> None:
        self.run_simple_query(
            "NIC Config",
            "nicconfig get description,dhcpserver,macaddress,ipaddress,defaultipgateway,ipsubnet",
            "IP configuration and DHCP details"
        )

    def show_ntdomain(self) -> None:
        self.run_simple_query(
            "NT Domain",
            "ntdomain get domaincontrolleraddress,caption,clientsitename,dcsitename,dnsforestname,domaincontrollername,domainname",
            "Domain controller and Active Directory related details"
        )

    def show_onboarddevice(self) -> None:
        self.run_simple_query(
            "Onboard Device",
            "onboarddevice get description",
            "Motherboard integrated device information"
        )

    def show_os(self) -> None:
        self.run_simple_query(
            "Operating System",
            "os get caption,version,countrycode,oslanguage,muilanguages",
            "Installed Windows operating system details"
        )

    def show_printer(self) -> None:
        self.run_simple_query(
            "Printer Information",
            "printer list status",
            "Installed printer overview"
        )

    def show_process(self) -> None:
        self.run_simple_query(
            "Process Information",
            "process get name",
            "Running process names"
        )

    def show_product(self) -> None:
        self.run_simple_query(
            "Installed Products",
            "product list brief",
            "Installed software inventory from WMIC product"
        )

    def show_service(self) -> None:
        self.run_simple_query(
            "Service Information",
            "service get name",
            "Windows service list"
        )

    def show_share(self) -> None:
        self.run_simple_query(
            "Share Information",
            "share get caption,name,path",
            "File share definitions"
        )

    def show_startup(self) -> None:
        self.run_simple_query(
            "Startup Information",
            "startup get",
            "Startup program entries"
        )

    def show_sysaccount(self) -> None:
        self.run_simple_query(
            "System Accounts",
            "sysaccount get",
            "Local and system account information"
        )

    def show_full_report(self) -> None:
        self.current_title = "Full Report"
        self.current_command = None
        self.set_status("Building full report...")

        sections = [
            ("BIOS Version", "bios get BIOSversion"),
            ("Computer System", "computersystem get name,username,manufacturer,model,domain"),
            ("CPU Information", "cpu get name,numberofcores,systemname"),
            ("CS Product", "csproduct get identifyingnumber,name,uuid"),
            ("Desktop Monitor", "desktopmonitor get screenheight,screenwidth,systemname"),
            ("Disk Drive", "diskdrive get model,name,size"),
            ("Local Disk", "logicaldisk get name,providername,systemname,description"),
            ("Net Use", "netuse get localname,username,remotepath"),
            ("NIC Information", "nic get adaptertype,macaddress,maxspeed,name,systemname,caption"),
            ("NIC Config", "nicconfig get description,dhcpserver,macaddress,ipaddress,defaultipgateway,ipsubnet"),
            ("NT Domain", "ntdomain get domaincontrolleraddress,caption,clientsitename,dcsitename,dnsforestname,domaincontrollername,domainname"),
            ("Onboard Device", "onboarddevice get description"),
            ("Operating System", "os get caption,version,countrycode,oslanguage,muilanguages"),
            ("Printer Information", "printer list status"),
            ("Process Information", "process get name"),
            ("Installed Products", "product list brief"),
            ("Service Information", "service get name"),
            ("Share Information", "share get caption,name,path"),
            ("Startup Information", "startup get"),
            ("System Accounts", "sysaccount get"),
        ]

        report_lines = []
        report_lines.append("SYSINSPECTOR FULL REPORT")
        report_lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report_lines.append("")

        for title, wmic_args in sections:
            report_lines.append("=" * 100)
            report_lines.append(title.upper())
            report_lines.append("-" * 100)
            report_lines.append(self.run_wmic(wmic_args).strip())
            report_lines.append("")

        self.append_output(
            "Full Report",
            "\n".join(report_lines),
            "Complete inventory-style report for the local machine"
        )
        self.set_status("Full report loaded")


def main() -> None:
    root = tk.Tk()
    app = SysInspectorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()