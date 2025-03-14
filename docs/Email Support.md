# Email Support

## Currently Supported Providers

HAController is believed to work with any standard IMAP and SMTP servers. These are email settings the code has been tested to work with. Feel free to add other major providers if you've verified they work for both sending and receiving.

### PurelyMail

_IMAP Host:_ imap.purelymail.com _Port:_ 993

_SMTP Host:_ smtp.purelymail.com _Port:_ 587

Note: PurelyMail IMAP is not successfully tested at this time but mail checking in general has not been tested for a while.

### Gmail

_IMAP Host:_ imap.gmail.com _Port:_ 993

_SMTP Host:_ smtp.gmail.com _Port:_ 587

Note: Gmail users must enable "less secure login" in their Google account settings. Alternately, if you have two-factor authentication enabled, you can use an app-specific password instead.

### Outlook.com

_IMAP Host:_ imap-mail.outlook.com _Port:_ 993

_SMTP Host:_ smtp-mail.outlook.com _Port:_ 587

Note: Outlook.com seems to pretty aggressively treat HAController as a spam sender. Not recommended.

### HostGator

Note: Your HostGator mail account should work with IMAP 993 and SMTP 587. Use the domains recommended by your HostGator account.
