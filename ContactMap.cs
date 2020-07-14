using CsvHelper.Configuration;

namespace ExchangeContacts2CSV
{
    class ContactMap : ClassMap<ContactModel>
    {
        public ContactMap()
        {
            Map(m => m.Title).Index(0).Name("Titel");
            Map(m => m.GivenName).Index(1).Name("Vorname");
            Map(m => m.Prefix).Index(2).Name("Präfix");
            Map(m => m.Surname).Index(3).Name("Nachname");
            Map(m => m.CompanyName).Index(4).Name("Unternehmen");
            Map(m => m.EMailAddress).Index(5).Name("E-mail");
            Map(m => m.Memo).Index(6).Name("Memo");
            Map(m => m.CellBusiness).Index(7).Name("Mobilfunk Business");
            Map(m => m.CellBusinessSpeedDial).Index(8).Name("Mobilfunk Business Speeddial");
            Map(m => m.CellBusinessDescription).Index(9).Name("Mobilfunk Business Beschreibung");
            Map(m => m.Work).Index(10).Name("Arbeit");
            Map(m => m.WorkSpeedDial).Index(11).Name("Arbeit Speeddial");
            Map(m => m.WorkDescription).Index(12).Name("Arbeit Beschreibung");
            Map(m => m.Fax).Index(13).Name("Fax");
            Map(m => m.FaxSpeedDial).Index(14).Name("Fax Speeddial");
            Map(m => m.FaxDescription).Index(15).Name("Fax Beschreibung");
            Map(m => m.CellPersonal).Index(16).Name("Mobilfunk Privat");
            Map(m => m.CellPersonalSpeedDial).Index(17).Name("Mobilfunk Privat Speeddial");
            Map(m => m.CellPersonalDescription).Index(18).Name("Mobilfunk Privat Beschreibung");
            Map(m => m.Home).Index(19).Name("Zu Hause");
            Map(m => m.HomeSpeedDial).Index(20).Name("Zu Hause Speeddial");
            Map(m => m.HomeDescription).Index(21).Name("Zu Hause Beschreibung");
            Map(m => m.Work2).Index(22).Name("Work2");
            Map(m => m.Work2SpeedDial).Index(23).Name("Work2 Speeddial");
            Map(m => m.Work2Description).Index(24).Name("Work2 Beschreibung");
            Map(m => m.Home2).Index(25).Name("Home2");
            Map(m => m.Home2SpeedDial).Index(26).Name("Home2 Speeddial");
            Map(m => m.Home2Description).Index(27).Name("Home2 Beschreibung");
        }
    }
}
