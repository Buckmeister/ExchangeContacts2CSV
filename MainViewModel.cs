using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Globalization;
using System.Security;
using System.Windows.Input;
using System.IO;
using System.Net;
using Microsoft.Exchange.WebServices.Data;
using CsvHelper;
using BCK;
using System.Text.RegularExpressions;

namespace ExchangeContacts2CSV
{
    class MainViewModel : ViewModelBase
    {
        private readonly ExchangeService exService;
        public MainViewModel()
        {
            exService = new ExchangeService(ExchangeVersion.Exchange2013_SP1); 
        }

        private ObservableCollection<ContactModel> _contacts = new ObservableCollection<ContactModel>();
        public ObservableCollection<ContactModel> Contacts
        {
            get { return _contacts; }

            set
            {
                _contacts = value;
                OnPropertyChanged(nameof(Contacts));
            }
        }

        public string Username
        {
            get { return Properties.Settings.Default.Username; }
            set
            {
                Properties.Settings.Default.Username = value;
                OnPropertyChanged(nameof(Username));
                HandleValidationError(IsValidEmailAddress(value), nameof(Username), "Keine gültige E-Mail-Adresse.");

                // reset all app data
                Contacts.Clear();
                ContactsFolderEntries.Clear();
                CurrentContactsFolder = null;
                PasswordBoxEnabled = true;
                LoggedIn = false;
                HasData = false;
            }
        }

        private bool IsValidEmailAddress(string emailaddress)
        {
            bool isValid;

            try
            {   
                Match match = Regex.Match(emailaddress, "^[a-zA-Z0-9_]+([-+.'][a-zA-Z0-9_]+)*@[a-zA-Z0-9]+([-.][a-zA-Z0-9]+)*\\.[a-zA-Z0-9]+([-.][a-zA-Z0-9]+)*$");
                if (match.Success)
                {
                    isValid = true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
            return isValid;
        }

        public SecureString SecurePassword
        {
            private get; set;
        }

        private ContactsFolder _currentContactsFolder;
        public ContactsFolder CurrentContactsFolder
        {
            get { return _currentContactsFolder; }
            set
            {
                _currentContactsFolder = value;
                OnPropertyChanged(nameof(CurrentContactsFolder));
                Contacts.Clear();
                HasData = false;
            }
        }

        private ObservableCollection<ContactsFolderEntry> _contactsFolderEntries = new ObservableCollection<ContactsFolderEntry>();
        public ObservableCollection<ContactsFolderEntry> ContactsFolderEntries
        {
            get { return _contactsFolderEntries; }
            set
            {
                _contactsFolderEntries = value;
                OnPropertyChanged(nameof(ContactsFolderEntries));
            }
        }

        public string PathToCsv
        {
            get { return Properties.Settings.Default.PathToCsv; }
            set
            {
                Properties.Settings.Default.PathToCsv = value;
                OnPropertyChanged(nameof(PathToCsv));
                HandleValidationError(IsValidDirectory(value), nameof(PathToCsv), "Keine existierendes Verzeichnis.");
            }
        }

        private bool IsValidDirectory(string filePath)
        {
            return Directory.Exists(Path.GetDirectoryName(filePath));
        }

        private bool _loginEnabled;
        public bool LoginEnabled
        {
            get { return _loginEnabled; }
            set
            {
                _loginEnabled = value;
                OnPropertyChanged(nameof(LoginEnabled));
            }
        }

        private IAsyncCommand _loginCommand;
        public IAsyncCommand LoginCommand
        {
            get
            {
                return _loginCommand ?? 
                    (
                    _loginCommand = new AsyncCommand(LoginExecuteAsync, LoginCanExecute, new ConsoleErrorHandler())
                    );
            }
        }

        private bool LoginCanExecute()
        {
            if (HasPropertyError(nameof(Username)))
            {
                LoginEnabled = false;
                return false;
            }

            if (LoggedIn)
            {
                LoginEnabled = false;
                return false;
            }

            if (IsBusy)
            {
                LoginEnabled = false;
                return false;
            }

            LoginEnabled = true;
            return true;
        }

        private async System.Threading.Tasks.Task LoginExecuteAsync()
        {
            try
            {
                IsBusy = true;
                LoggedIn = await FindExchangeContactsFoldersAsync();
                PasswordBoxEnabled = !LoggedIn;
            }
            catch (Exception ex)
            {
                Console.WriteLine("EWS Fehler: " + ex.Message);
                if (DebugMode) Console.WriteLine(ex.ToString());
            }
            finally
            {
                IsBusy = false;
            }
        }

        private bool _isBusy;
        public bool IsBusy
        {
            get { return _isBusy; }
            set
            {
                _isBusy = value;
                OnPropertyChanged(nameof(IsBusy));
                
                CommandManager.InvalidateRequerySuggested();
            }
        }

        private bool _loggedIn;
        public bool LoggedIn
        {
            get { return _loggedIn; }
            set
            {
                _loggedIn = value;
                OnPropertyChanged(nameof(LoggedIn));
                
                CommandManager.InvalidateRequerySuggested();
            }
        }

        private bool _hasData;
        public bool HasData
        {
            get { return _hasData; }
            set
            {
                _hasData = value;
                OnPropertyChanged(nameof(HasData));

                CommandManager.InvalidateRequerySuggested();
            }
        }

        private bool _debugMode;
        public bool DebugMode
        {
            get { return _debugMode; }
            set
            {
                _debugMode = value;
                OnPropertyChanged(nameof(DebugMode));
            }
        }

        private bool _passwordBoxEnabled = true;
        public bool PasswordBoxEnabled
        {
            get { return _passwordBoxEnabled; }
            set
            {
                _passwordBoxEnabled = value;
                OnPropertyChanged(nameof(PasswordBoxEnabled));
            }
        }


        private bool _importEnabled;
        public bool ImportEnabled
        {
            get { return _importEnabled; }
            set
            {
                _importEnabled = value;
                OnPropertyChanged(nameof(ImportEnabled));
            }
        }

        private IAsyncCommand _importCommand;
        public IAsyncCommand ImportCommand
        {
            get
            {
                return _importCommand ??
                    (
                    _importCommand = new AsyncCommand(ImportExecuteAsync, ImportCanExecute, new ConsoleErrorHandler())
                    );
            }
        }

        public bool ImportCanExecute()
        {
            if (!LoggedIn)
            {
                ImportEnabled = false;
                return false;
            }

            if (HasData)
            {
                ImportEnabled = false;
                return false;
            }

            if (IsBusy)
            {
                ImportEnabled = false;
                return false;
            }

            ImportEnabled = true;
            return true;
        }

        public async System.Threading.Tasks.Task ImportExecuteAsync()
        {
            try
            {
                IsBusy = true;
                HasData = await ImportExchangeContactsAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Import Fehler: " + ex.Message);
                if (DebugMode) Console.WriteLine(ex.ToString());
                
            }
            finally
            {
                IsBusy = false;
            }
        }

        private bool _exportEnabled;
        public bool ExportEnabled
        {
            get { return _exportEnabled; }
            set
            {
                _exportEnabled = value;
                OnPropertyChanged(nameof(ExportEnabled));
            }
        }

        private IAsyncCommand _exportCommand;
        public IAsyncCommand ExportCommand
        {
            get
            {
                return _exportCommand ??
                    (
                    _exportCommand = new AsyncCommand(ExportExecuteAsync, ExportCanExecute, new ConsoleErrorHandler())
                    );
            }
        }

        public bool ExportCanExecute()
        {
            if (HasPropertyError(nameof(PathToCsv)))
            {
                ExportEnabled = false;
                return false;
            }

            if (!LoggedIn)
            {
                ExportEnabled = false;
                return false;
            }

            if (!HasData)
            {
                ExportEnabled = false;
                return false;
            }

            ExportEnabled = true;
            return true;
        }

        public async System.Threading.Tasks.Task ExportExecuteAsync()
        {
            try
            {
                IsBusy = true;
                await WriteCsvRecordsAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Export Fehler: " + ex.Message);
                if (DebugMode) Console.WriteLine(ex.ToString());
            }
            finally
            {
                IsBusy = false;
            }
        }

        private async Task<bool> FindExchangeContactsFoldersAsync()
        {
            try // Autodiscover
            {
                exService.TraceListener = new EwsTraceListener();
                exService.TraceFlags = TraceFlags.AutodiscoverConfiguration |
                                        TraceFlags.AutodiscoverRequest |
                                        TraceFlags.AutodiscoverResponse;

                exService.TraceEnabled = DebugMode;
                exService.Credentials = new NetworkCredential(Username, SecurePassword);

                await System.Threading.Tasks.Task.Run(() =>
                {
                    exService.AutodiscoverUrl(Username, redirectionUrlValidationCallback);
                });
            }
            catch (ServiceRequestException ex)
            {
                Console.WriteLine("Exchange Fehler: " + ex.Message);
                if (DebugMode) Console.WriteLine(ex.ToString());

                LoggedIn = false;
                return false;
            }
            catch (ServiceResponseException ex)
            {
                Console.WriteLine("Exchange Fehler: " + ex.Message);
                if (DebugMode) Console.WriteLine(ex.ToString());

                LoggedIn = false;
                return false;
            }
            catch (AutodiscoverLocalException ex)
            {
                Console.WriteLine("Exchange Fehler: " + ex.Message);
                if (DebugMode) Console.WriteLine(ex.ToString());

                LoggedIn = false;
                return false;
            }

            Console.WriteLine("EWS Endpoint: {0}", exService.Url);

            try // Find local contacts folders in user's mailbox
            {
                ContactsFolder contactsFolder = await System.Threading.Tasks.Task.Run(() =>
                {
                   return ContactsFolder.Bind(exService, WellKnownFolderName.Contacts);
                });

                ContactsFolderEntries.Add(new ContactsFolderEntry
                {
                    Display = contactsFolder.DisplayName,
                    ContactsFolder = contactsFolder
                });

                CurrentContactsFolder = contactsFolder;

                
                FindFoldersResults allMailboxContactsFolders = await System.Threading.Tasks.Task.Run(() =>
                {
                    var contactsFolderFilter = new SearchFilter.IsEqualTo(
                                                        FolderSchema.FolderClass,
                                                        CurrentContactsFolder.FolderClass
                                                        );

                    return exService.FindFolders(
                            WellKnownFolderName.MsgFolderRoot,
                            contactsFolderFilter,
                            new FolderView(int.MaxValue)
                            { Traversal = FolderTraversal.Deep }
                            );
                });

                foreach (var folder in allMailboxContactsFolders)
                {
                    if (folder.Id.UniqueId == CurrentContactsFolder.Id.UniqueId) continue;
                    if (folder.DisplayName == "ExternalContacts") continue;
                    if (folder.DisplayName == "PersonMetadata") continue;

                    ContactsFolderEntries.Add(new ContactsFolderEntry
                    {
                        Display = folder.DisplayName,
                        ContactsFolder = folder as ContactsFolder
                    });
                }

                
            }
            catch (ServiceRequestException ex)
            {
                Console.WriteLine("Exchange Fehler: " + ex.Message);
                if (DebugMode) Console.WriteLine(ex.ToString());

                LoggedIn = false;
                return false;
            }
            catch (ServiceResponseException ex)
            {
                Console.WriteLine("Exchange Fehler: " + ex.Message);
                if (DebugMode) Console.WriteLine(ex.ToString());

                LoggedIn = false;
                return false;
            }
            
            try // Find contacts folders in public folder tree
            {
                FindFoldersResults allPublicContactsFolders = await System.Threading.Tasks.Task.Run(() =>
                {
                    var contactsFolderFilter = new SearchFilter.IsEqualTo(
                                                        FolderSchema.FolderClass,
                                                        CurrentContactsFolder.FolderClass
                                                        );

                    return exService.FindFolders(
                            WellKnownFolderName.PublicFoldersRoot,
                            contactsFolderFilter,
                            new FolderView(int.MaxValue)
                            );
                });

                foreach (var folder in allPublicContactsFolders)
                {
                    ContactsFolderEntries.Add(new ContactsFolderEntry
                    {
                        Display = folder.DisplayName,
                        ContactsFolder = folder as ContactsFolder
                    });
                }
            }
            catch (ServiceRequestException ex)
            {
                Console.WriteLine("Exchange Fehler: " + ex.Message);
                if (DebugMode) Console.WriteLine(ex.ToString());

                LoggedIn = false;
                return false;
            }
            catch (ServiceResponseException ex)
            {
                Console.WriteLine("Exchange Fehler: " + ex.Message);
                if (DebugMode) Console.WriteLine(ex.ToString());

                // Do nothing, if public folder search doesn't return successfully.
            }
            

            Console.WriteLine("Es wurden {0} Kontakte-Ordner identifiziert.", ContactsFolderEntries.Count);

            return true;
        }

        private static bool redirectionUrlValidationCallback(string redirectionUrl)
        {
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate if the connection is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        private async Task<bool> ImportExchangeContactsAsync()
        {
            try
            {
                if (CurrentContactsFolder.TotalCount == 0)
                {
                    Console.WriteLine($"Ordner '{CurrentContactsFolder.DisplayName}' enthält keine Elemente.");
                    return false;
                }

                FindItemsResults<Item> foundItems = await System.Threading.Tasks.Task.Run(() => {
                    Console.WriteLine($"Ordner '{CurrentContactsFolder.DisplayName}' wird durchsucht.");
                    ItemView itemView = new ItemView(CurrentContactsFolder.TotalCount);
                    return exService.FindItems(CurrentContactsFolder.Id, itemView);
                    });

                
                Console.WriteLine("Es wurde{0} {1} Objekt{2} im ausgewählten Ordner gefunden.",
                                    foundItems.TotalCount == 1 ? "" : "n",
                                    foundItems.TotalCount, 
                                    foundItems.TotalCount==1 ? "" : "e");

                foreach (Item item in foundItems)
                {
                    if (item is Contact)
                    {
                        Contact exContact = item as Contact;
                        Console.WriteLine("");
                        Console.WriteLine("Kontakt: " + exContact.DisplayName);

                        ContactModel contact = new ContactModel
                        {
                            Title = exContact.JobTitle,
                            GivenName = exContact.GivenName,
                            Surname = exContact.Surname,
                            CompanyName = exContact.CompanyName
                        };

                        if (exContact.EmailAddresses.TryGetValue(EmailAddressKey.EmailAddress1, out EmailAddress emailAdress))
                        {
                            if (IsValidEmailAddress(emailAdress.Address))
                            {
                                contact.EMailAddress = emailAdress.Address;
                            }
                        }

                        string phoneNumber;

                        phoneNumber = string.Empty;
                        exContact.PhoneNumbers.TryGetValue(PhoneNumberKey.MobilePhone, out phoneNumber);
                        if (!string.IsNullOrEmpty(phoneNumber) && !IsValidPhoneNumber(phoneNumber)) continue;
                        contact.CellBusiness = ReFormatPhoneNumber(phoneNumber);

                        phoneNumber = string.Empty;
                        exContact.PhoneNumbers.TryGetValue(PhoneNumberKey.BusinessPhone, out phoneNumber);
                        if (!string.IsNullOrEmpty(phoneNumber) && !IsValidPhoneNumber(phoneNumber)) continue;
                        contact.Work = ReFormatPhoneNumber(phoneNumber);

                        phoneNumber = string.Empty;
                        exContact.PhoneNumbers.TryGetValue(PhoneNumberKey.BusinessFax, out phoneNumber);
                        if (!string.IsNullOrEmpty(phoneNumber) && !IsValidPhoneNumber(phoneNumber)) continue;
                        contact.Fax = ReFormatPhoneNumber(phoneNumber);

                        phoneNumber = string.Empty;
                        exContact.PhoneNumbers.TryGetValue(PhoneNumberKey.OtherTelephone, out phoneNumber);
                        if (!string.IsNullOrEmpty(phoneNumber) && !IsValidPhoneNumber(phoneNumber)) continue;
                        contact.CellPersonal = ReFormatPhoneNumber(phoneNumber);

                        phoneNumber = string.Empty;
                        exContact.PhoneNumbers.TryGetValue(PhoneNumberKey.HomePhone, out phoneNumber);
                        if (!string.IsNullOrEmpty(phoneNumber) && !IsValidPhoneNumber(phoneNumber)) continue;
                        contact.Home = ReFormatPhoneNumber(phoneNumber);

                        phoneNumber = string.Empty;
                        exContact.PhoneNumbers.TryGetValue(PhoneNumberKey.BusinessPhone2, out phoneNumber);
                        if (!string.IsNullOrEmpty(phoneNumber) && !IsValidPhoneNumber(phoneNumber)) continue;
                        contact.Work2 = ReFormatPhoneNumber(phoneNumber);

                        phoneNumber = string.Empty;
                        exContact.PhoneNumbers.TryGetValue(PhoneNumberKey.HomePhone2, out phoneNumber);
                        if (!string.IsNullOrEmpty(phoneNumber) && !IsValidPhoneNumber(phoneNumber)) continue;
                        contact.Home2 = ReFormatPhoneNumber(phoneNumber);

                        Contacts.Add(contact);  
                    }
                }
                Console.WriteLine("Es wurde{0} {1} Kontakt{2} erfolgreich importiert.",
                                    Contacts.Count == 1 ? "" : "n",
                                    Contacts.Count, 
                                    Contacts.Count==1?"":"e");
            }
            catch (ServiceRequestException ex)
            {
                Console.WriteLine("Exchange Fehler: " + ex.Message);
                if (DebugMode) Console.WriteLine(ex.ToString());

                LoggedIn = false;
                return false;
            }
            catch (ServiceResponseException ex)
            {
                Console.WriteLine("Exchange Fehler: " + ex.Message);
                if (DebugMode) Console.WriteLine(ex.ToString());

                LoggedIn = false;
                return false;
            }
            return true;
        }

        private bool IsValidPhoneNumber(string phoneNumber)
        {
            phoneNumber = Regex.Replace(phoneNumber, "[^+0-9]", "");
            phoneNumber = Regex.Replace(phoneNumber, "^00", "+");

            int lengthOfPhoneNumber = phoneNumber.Length;

            // if a country code is used, we don't count the plus sign. 
            // so we can subtract 1 from length.
            if (phoneNumber.StartsWith("+")) lengthOfPhoneNumber -= 1;

            // if a german country code is used (e.g. +49) it translates into a zero.
            // so we can substract 1 again, in that case.
            if (phoneNumber.StartsWith("+49")) lengthOfPhoneNumber -= 1;

            // the remaining number... 
            // ...has to be at least 6 digits...
            if (lengthOfPhoneNumber < 6)
            {
                Console.WriteLine($"Rufnummer zu kurz: {phoneNumber} hat {lengthOfPhoneNumber} Stellen.");
                Console.WriteLine("Kontakt wird übersprungen.");

                return false;
            }

            // ...but not more than 15 digits.
            if (lengthOfPhoneNumber > 15) 
            {
                Console.WriteLine($"Rufnummer zu lang: {phoneNumber} hat {lengthOfPhoneNumber} Stellen.");
                Console.WriteLine("Kontakt wird übersprungen.");

                return false; 
            }

            return true;
        }

        private string ReFormatPhoneNumber(string phoneNumber)
        {
            try
            {
                // Removing all but digits, "+", "-", "()" and the silly (and totally wrong) "(0)" bs.
                phoneNumber = Regex.Replace(phoneNumber, "[^0-9 +-/()]|\\(0\\)", "");

                // Replace leading double zero with plus sign.
                phoneNumber = Regex.Replace(phoneNumber, "^00", "+");
            }
            catch (Exception)
            {
                phoneNumber = string.Empty;
            }
            return phoneNumber;
        }

        private async System.Threading.Tasks.Task WriteCsvRecordsAsync()
        {
            await System.Threading.Tasks.Task.Run(() => {
                using (var writer = new StreamWriter(PathToCsv, false, System.Text.Encoding.GetEncoding("ISO-8859-1")))
                using (var csv = new CsvWriter(writer, CultureInfo.GetCultureInfo("de-DE")))
                {
                    csv.Configuration.RegisterClassMap<ContactMap>();
                    csv.WriteRecords(Contacts);
                }
                Console.WriteLine("CSV-Datei wurde erstellt:");
                Console.WriteLine(Path.GetFullPath(PathToCsv));
            });
        }
    }
}
