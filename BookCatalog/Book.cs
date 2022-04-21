using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BookCatalog
{
    internal class Book
    {
        public string title { get; set; }
        public string author { get; set; }
        public string isbn { get; set; }

        public Book(string title, string author, string isbn)
        {
            this.title = title;
            this.author = author;  
            this.isbn = isbn;
        }

        
    }
}
