using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace ADUMP
{
    internal static class DomainConfiguration
    {
        public static bool IsValid(string domain) => domain.Length <= 253 &&
            domain.Split('.').All(label => Regex.IsMatch(label, @"\A[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?\z"));

        public static string DistinguishedName(string domain)
        {
            if (!IsValid(domain)) throw new ArgumentException("Nom de domaine DNS invalide.", nameof(domain));
            return string.Join(",", domain.Split('.').Select(label => "DC=" + label));
        }

        public static string Delimiter(string setting) => setting switch
        {
            "semicolon" => ";",
            "tab" => "\t",
            _ => ","
        };
    }
}
