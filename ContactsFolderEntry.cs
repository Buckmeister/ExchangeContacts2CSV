using Microsoft.Exchange.WebServices.Data;

namespace ExchangeContacts2CSV
{
    class ContactsFolderEntry
    {
        public ContactsFolderEntry()
        {

        }

        public string Display { get; set; }
        public ContactsFolder ContactsFolder { get; set; }
    }
}
