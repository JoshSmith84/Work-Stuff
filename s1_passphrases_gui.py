#! python3

import csv
import requests
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import sys
import datetime
import os

class LabelInput(tk.Frame):
    """A widget containing a label and input together."""

    def __init__(
            self, parent, label, var, input_class=ttk.Entry,
            input_args=None, label_args=None, **kwargs
    ):
        super().__init__(parent, **kwargs)
        input_args = input_args or {}
        label_args = label_args or {}
        self.variable = var
        self.variable.label_widget = self

        if input_class in (ttk.Checkbutton, ttk.Button):
            input_args["text"] = label
        else:
            self.label = ttk.Label(self, text=label, **label_args)
            self.label.grid(row=0, column=0, sticky=(tk.W + tk.E))

        if input_class in (
            ttk.Checkbutton, ttk.Button, ttk.Radiobutton
        ):
            input_args["variable"] = self.variable
        else:
            input_args["textvariable"] = self.variable

        if input_class == ttk.Radiobutton:
            self.input = tk.Frame(self)
            for v in input_args.pop('values', []):
                button = ttk.Radiobutton(
                    self.input, value=v, text=v, **input_args
                )
                button.pack(
                    side=tk.LEFT, ipadx=10,
                    ipady=2, expand=True, fill='x'
                )
        else:
            self.input = input_class(self, **input_args)

        self.input.grid(row=1, column=0, sticky=(tk.E + tk.W))
        self.columnconfigure(0, weight=1)

    def grid(self, sticky=(tk.E + tk.W), **kwargs):
        """Override grid to add default sticky values"""
        super().grid(sticky=sticky, **kwargs)


class AppPage(ttk.Frame):
    """Application page class from which all other pages will inherit."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._vars = {}

    def _add_frame(self, label, cols=2):

        frame = ttk.LabelFrame(self, text=label)
        frame.grid(sticky=tk.W + tk.E)
        for i in range(cols):
            frame.columnconfigure(i, weight=1)
        return frame

    def get(self):
        data = dict()
        for key, variable in self._vars.items():
            try:
                data[key] = variable.get()
            except tk.TclError:
                message = f'Error in field: {key}.'
                raise ValueError(message)
        return data

class MainPage(AppPage):
    """Main Page to select options, change folder, and run"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._vars = {'Portal': tk.StringVar(None, 'Ntiva'),
                      'Client Name': tk.StringVar(None, None),
                      'API': tk.StringVar(None, None),
                      'SiteID': tk.StringVar(None, None),
                      }

        self.output_folder = os.path.join(os.path.expanduser("~"), "Documents")
        self.err_file = os.path.join(self.output_folder, "s1-passphrase-error.txt")
        self.portal_url = ''
        self.result_list = []
        self.fields = ["computerName", "domain", "passphrase", "portalStatus"]

        # Initialize Main Page GUI
        partner_default = self._add_frame(
            'Select Partner Portal'
        )
        client_default = self._add_frame(
            'Client Name'
        )
        api_default = self._add_frame('API Key')
        site_default = self._add_frame('SiteID')
        buttons = self._add_frame('')

        LabelInput(partner_default, '', input_class=ttk.Radiobutton,
                   var=self._vars['Portal'],
                   input_args={'values': ['Ntiva', 'TPG']}
                   ).grid(row=0, column=0, sticky=(tk.W + tk.E)
                          )

        LabelInput(client_default, '',
                   var=self._vars['Client Name'],
                   ).grid(row=1, column=0, sticky=(tk.W + tk.E)
                          )

        LabelInput(api_default, '',
                   var=self._vars['API'],
                   ).grid(row=2, column=0, sticky=(tk.W + tk.E), columnspan=8
                          )

        LabelInput(site_default, '',
                   var=self._vars['SiteID'],
                   ).grid(row=3, column=0, sticky=(tk.W + tk.E)
                          )

        self.run_button = tk.Button(buttons, text='Run',
                                    command=self._on_run
                                    )
        self.run_button.grid(row=0, column=1, sticky='ew')

        self.select_target = tk.Button(buttons, text='Change Folder',
                                       command=self._on_target
                                       )
        self.select_target.grid(row=0, column=0, sticky='ew')

        self.quit_button = tk.Button(
            buttons,
            text='Quit',
            command=self._on_quit
        )
        self.quit_button.grid(row=4, column=0, sticky='ew')

        self.status = tk.StringVar(
            None, 'Status: '
                  'Please change output folder or run with defaults\n(default is your Documents folder)...'
        )
        ttk.Label(
            self, textvariable=self.status, wraplength=225, justify='left'
        ).grid(sticky=(tk.W + tk.E), row=5, padx=10)


    def pull_results(self, base_url, api_key: str, site_id: str, decom):
        """
        Method to pull S1 passphrases from S1 portal.

        param: base_url: Base URL of S1 portal. Only options right now are Ntiva and TPG
        param: api_key: API key from S1 portal. The key must belong to a service account
        that at least has admin access to the Site ID being pulled
        param: site_id: Site ID of the client site in S1 portal
        return: None
        """

        s1_headers = {
            "Accept": "application/json",
            "User-Agent": "vz/s1_agent_passphrases_v1.0",
            "Content-Type": "application/json",
        }
        s1_passphrase_api_endpoint = '/web/api/v2.1/agents/passphrases'

        limit = 200
        headers = {**s1_headers, "Authorization": "ApiToken " + api_key}
        url = base_url + s1_passphrase_api_endpoint
        params = {"limit": limit,
                  "siteIds": site_id,
                  "isDecommissioned": decom}
        next_cursor = None
        done, errored = False, False

        while not (done or errored):
            if next_cursor:
                params = {**params, "cursor": next_cursor}

            response = requests.get(url, headers=headers, params=params, )

            if response.status_code != requests.codes.ok:
                errored = True
                if response.status_code == 400:
                    self.status.set(f"error getting data: {response.status_code}. Check your SiteID and try again.")
                elif response.status_code == 401:
                    self.status.set(f"error getting data: {response.status_code}. Check your API key and try again")
                else:
                    self.status.set(f"error getting data: {response.status_code}. "
                                    f"See 's1-passphrase-error file in output folder for more info")
                    self.log_error(self.err_file, str(response.headers))
            else:
                data = response.json()

                next_cursor = data["pagination"]["nextCursor"]
                if next_cursor is None:
                    done = True

                if "data" in data:
                    for item in data["data"]:
                        yield {k: item[k] for k in self.fields if k in item}
                    self.status.set(f'Processing Complete. See {self.output_folder} for results.')
                else:
                    errored = True
                    self.status.set("errors present")
                    self.log_error(self.err_file, data["errors"])

                del data

    def _on_run(self):
        """Command to grab the Api Key and SiteID variables from the form and request from S1 portal"""

        if self._vars['Portal'].get() == 'Ntiva':
            self.portal_url = 'https://usea1-cw04mdr.sentinelone.net'
        else:
            self.portal_url = 'https://usea1-pax8-03.sentinelone.net'

        for decom_answer in ['no','yes']:
            for line in self.pull_results(
                    self.portal_url,
                    self._vars['API'].get().strip(),
                    self._vars['SiteID'].get().strip(),
                    decom_answer
            ):
                if decom_answer == 'yes':
                    line['portalStatus']='decommissioned'
                else:
                    line['portalStatus']='active'
                self.result_list.append(line)

        output_file = os.path.join(self.output_folder, f's1-passphrases-{
        self._vars['Client Name'].get().strip().replace(" ", "_")
        }.csv')

        with open(output_file, "w", newline="") as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=self.fields)
            writer.writeheader()

            for row in self.result_list:
                writer.writerow(row)


    def _on_target(self):
        """Command to change the output folder"""

        ch_folder_diag = tk.Tk()
        ch_folder_diag.overrideredirect(True)
        ch_folder_diag.attributes('-alpha', 0)
        ch_folder_diag.title('Choose output folder...')
        self.output_folder = filedialog.askdirectory(
            title='Choose output folder...',
        )
        self.err_file = os.path.join(self.output_folder, "s1-passphrase-error.txt")
        ch_folder_diag.destroy()

        self.status.set(f'Output Folder set to: \n{self.output_folder}. '
                        f'\nChoose Run to continue...'
                        )

    @staticmethod
    def log_error(err_file, message) -> None:
        """Simple method for opening passed txt file
        and appending message

        :param err_file: txt file
        :param message: string to append.
        This method will prepend current date and time
        """

        with open(err_file, 'a') as f:
            f.write(f'Error_{datetime.datetime.now()}_{message}\n')

    @staticmethod
    def _on_quit():
        """Command to exit program"""
        sys.exit()



class Application(tk.Tk):
    """Application root window"""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.m_page = ''
        self.main_label = ''
        self.title(" S1 Passphrases v1.0")
        self.minsize(400, 350)
        self.main_page()

    def main_page(self):
        self.m_page = MainPage(self)
        self.main_label = ttk.Label(
            self,
            text="S1 Passphrases",
            font=("TKDefaultFont", 14))
        self.main_label.grid(row=0)
        self.m_page.grid(row=1, padx=10, sticky=(tk.W + tk.E))


if __name__ == "__main__":
    app = Application()
    app.grid_columnconfigure(0, weight=1)
    app.mainloop()