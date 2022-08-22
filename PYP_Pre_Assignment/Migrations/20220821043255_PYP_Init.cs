using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PYP_Pre_Assignment.Migrations
{
    public partial class PYP_Init : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "XLSFiles",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Segment = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Country = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Product = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    DiscountBand = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    UnitsSold = table.Column<double>(type: "float", nullable: false),
                    ManufacturingPrice = table.Column<double>(type: "float", nullable: false),
                    SalePrice = table.Column<double>(type: "float", nullable: false),
                    GrossSales = table.Column<double>(type: "float", nullable: false),
                    Discounts = table.Column<double>(type: "float", nullable: false),
                    Sales = table.Column<double>(type: "float", nullable: false),
                    COGS = table.Column<double>(type: "float", nullable: false),
                    Profit = table.Column<double>(type: "float", nullable: false),
                    Date = table.Column<DateTime>(type: "datetime2", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_XLSFiles", x => x.Id);
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "XLSFiles");
        }
    }
}
