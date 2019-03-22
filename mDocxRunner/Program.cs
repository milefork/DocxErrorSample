using System;
using System.Collections.Generic;
using System.IO;
using Xceed.Words.NET;

namespace mDocxRunner
{
   class Program
   {
      static void Main(string[] args)
      {

         var dict = new Dictionary<string, string>
         {
            {"{ServiceNr}","" },
            {"{FirmaName}","asdasd" },
            {"{FirmaStrasse}","sdfsdfsdf" },
            {"{FirmaPlzStadt}","234234234" },
            {"{FirmaTel}","" },
            {"{FirmaFax}","" },
            {"{FirmaEmail}","" },
            {"{KDNummer}","" },
            {"{KDName}","" },
            {"{KDStrasse}",""},
            {"{KDPlzStadt}",""},
            {"{KDTel}",""},
            {"{KDEmail}",""},
            {"{SVTicketID}","asdasda"},
            {"{SVMaxKosten}","500,00 EUR"},
            {"{GERBesch}","sdfsdfs"},
            {"{GERZugang}","sdfsdfs"},
            {"{AUFBesch}","asdasdasdas"},
            {"{FirmaAngenommenMitarbeiter}","xyz"},
            {"{FirmaAbgegebenMitarbeiter}",""},
            {"{KDVonMitarbeiter}","sdfasdf"},
            {"{KDZuMitarbeiter}",""},
            {"{Datum}","22.03.2019"},
            {"{SVFertigBisBeschreibung}","" }
         };

         var path = Path.GetFullPath(Path.Combine($"{System.AppDomain.CurrentDomain.BaseDirectory}", "../../VorlageServiceAuftrag.docx"));
         var doc = DocX.Load(path);

         foreach (var dictKey in dict.Keys)
         {
            if (string.IsNullOrWhiteSpace(dict[dictKey]))
            {
               doc.ReplaceText(dictKey, "");
            }
            else
            {
               doc.ReplaceText(dictKey, dict[dictKey]);
            }
         }

         doc.Save();
      }
   }
}
