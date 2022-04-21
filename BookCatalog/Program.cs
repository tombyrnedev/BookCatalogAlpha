using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Collections.Generic;
using System.Text.Json;
using BookCatalog;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

List<Book> books = new();
Console.WriteLine("Welcome to Simple Book Catalog V0.1");
Console.WriteLine("Scan ISBNs to add books (or type)");
Console.WriteLine("Type 0 to end program and generate Excel File");
Console.WriteLine("At the moment the file is saved as BookCat in your documents folder");
Console.WriteLine("AT THE MOMENT THERE IS NO MEMORY SO YOU NEED TO COPY the Excel File away from documents and do a manual merge");

MainAsync(args).GetAwaiter().GetResult();

async Task ProcessRepositories(String s)
{
    try
    {

        HttpClient client = new();
        HttpResponseMessage response = await client.GetAsync("https://www.googleapis.com/books/v1/volumes?q=isbn:" + s);
        response.EnsureSuccessStatusCode();
        string responseBody = await response.Content.ReadAsStringAsync();
        
        //Console.WriteLine(responseBody);
        DisplayResult(responseBody);

    }
    catch (HttpRequestException e)
    {
        Console.WriteLine("\nException Caught!");
        Console.WriteLine("Message :{0} ", e.Message);
    }
}

async Task MainAsync(string[] args)
{
    
    string ib = "";
    while(ib != "0")
    {
        Console.Write("-> ");
        ib = Console.ReadLine();
        await ProcessRepositories(ib);
    }
    
    
    DisplayBooks();
}

void DisplayResult(string response){
    //Do something with the JSON
    //Console.WriteLine(response);
    try
    {
        JObject result = JObject.Parse(response);
        if (result != null)
        {
        
            JToken token = result.GetValue("items");
            try
            {
                String title = (string)token[0]["volumeInfo"]["title"];
                String author = (string)token[0]["volumeInfo"]["authors"][0];
                String isbn = (string)token[0]["volumeInfo"]["industryIdentifiers"][0]["identifier"];

                Book b = new Book(title, author, isbn);
                books.Add(b);
                Console.WriteLine("Added {0}", title);
            }
            catch (Exception e) {
                Console.WriteLine("Not a book or end of program requested");
            }

            
            
        }
        else
        {
            Console.WriteLine("No info");
        }
    }catch (Exception e)
    {
        Console.WriteLine("\nException Caught!");
        Console.WriteLine("Message :{0} ", e.Message);
    }
   
    
}

void DisplayBooks()
{
    Console.WriteLine("Books added this session");
    foreach(Book book in books)
    {
        Console.WriteLine(book.title);
        Console.WriteLine(book.author);
        Console.WriteLine(book.isbn);
        Console.WriteLine();
    }
    AddBooksToExcel();
}

void AddBooksToExcel()
{
    Console.WriteLine("Adding Books to Excel please wait");
    Excel.Application xlApp = new
    Microsoft.Office.Interop.Excel.Application();
    var xlwb = xlApp.Workbooks.Add();
    var xlws = (Excel.Worksheet)xlwb.Worksheets.get_Item(1);
    xlws.Cells[1, 1] = "Title";
    xlws.Cells[1, 2] = "Author";
    xlws.Cells[1, 3] = "ISBN";

    int r = 2;

    foreach (Book book in books)
    {
        xlws.Cells[r , 1] = book.title;
        xlws.Cells [r , 2] = book.author;
        xlws.Cells[r,3] = book.isbn;
        r++;
    }

    xlwb.SaveAs("BookCat");
    xlwb.Close();
    Marshal.ReleaseComObject(xlws);
    Marshal.ReleaseComObject(xlwb);
    Marshal.ReleaseComObject(xlApp);
    xlws = null;
    xlwb = null;
    xlApp = null;

    Console.WriteLine("Press Enter to end program\nRemember the file is BookCat in Documents");
    Console.ReadLine();


}
