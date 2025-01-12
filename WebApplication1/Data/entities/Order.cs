using System.ComponentModel.DataAnnotations;

namespace ExcelWithDotNet9.Web.Data.entities;

public class Order
{

    [Key]
    public int Id { get; set; }

    public string Sum { get; set; }

}