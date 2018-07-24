# Email Support

## Currently Supported Providers

HAController is believed to work with any standard IMAP and SMTP servers. It's POP3 support is questionable. These are email settings the code has been tested to work with. Feel free to add other major providers if you've verified they work for both sending and receiving.

### Gmail

_IMAP Host:_ imap.gmail.com _Port:_ 993

_POP3 Host:_ pop.gmail.com _Port:_ 995

_SMTP Host:_ smtp.gmail.com _Port:_ 587

Note: Gmail users must enable "less secure login" in their Google account settings. Alternately, if you have two-factor authentication enabled, you can use an app-specific password instead. Note that you must enable POP3 in your Gmail settings as well if you are using POP3.

### Outlook.com

_IMAP Host:_ imap-mail.outlook.com _Port:_ 993

_SMTP Host:_ smtp-mail.outlook.com _Port:_ 587

Note: Outlook.com seems to pretty aggressively treat HAController as a spam sender. Not recommended. HAC's POP3 client does not appear to work correctly with Outlook.com at this time.

### HostGator

Note: Your HostGator mail account should work with IMAP 993 and SMTP 587. POP isn't tested. Use the domains recommended by your HostGator account.