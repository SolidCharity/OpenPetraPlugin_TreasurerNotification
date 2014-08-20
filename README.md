# Plugin TreasurerNotification

This is a plugin for [www.openpetra.org](http://www.openpetra.org).

## Functionality

* You can send letters and emails to the treasurer of a worker. If an email address is stored, the email delivery is prefered over the normal letter.
* The message contains an overview of the totals of incoming donations for the specific worker.
* The treasurer is someone in the sending church, connected via partner relationship TREASURER to the worker.
* The email and letter text is defined in an HTML template.

## Dependencies

* There are no dependancies.

## Installation

Please copy this directory to your OpenPetra working directory, to csharp\ICT\Petra\Plugins, or include it like this, if you are using git anyway:

    git submodule add https://github.com/SolidCharity/OpenPetraPlugin_TreasurerNotification csharp/ICT/Petra/Plugins/TreasurerNotification

and then run

    nant generateSolution

Please check the config directory for changes to your config files.

## License

This plugin is licensed under the GPL v3.