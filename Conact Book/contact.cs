

namespace Conact_Book
{
    internal class contact
    {
        public string Name { get; private set; }
        public string Surname { get; private set; }
        public string Address { get; private set; }
        public string CellPhone { get; private set; }

        
        public contact(string name, string surname, string address, string cellPhone)
        {
            Name = name;
            Surname = surname;
            Address = address;
            CellPhone = cellPhone;
        }
    }
}
