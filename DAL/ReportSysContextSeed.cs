using System;
using System.IO;
using OfficeOpenXml;
using System.Data;
using ReportSys.DAL.Entities;
using Microsoft.AspNetCore.Identity;

namespace ReportSys.DAL
{
    public class ReportSysContextSeed
    {

        

       

        static DataTable LoadExcelFile(string filePath)
        {
            FileInfo fileInfo = new FileInfo(filePath);
            DataTable dataTable = new DataTable();

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Используем первый лист

                // Добавляем колонки
                foreach (var firstRowCell in worksheet.Cells[4, 1, 4, worksheet.Dimension.End.Column])
                {
                    dataTable.Columns.Add(firstRowCell.Text);
                }

                // Добавляем строки
                for (int rowNum = 5; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                {
                    var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                    DataRow row = dataTable.NewRow();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                    dataTable.Rows.Add(row);
                }
            }

            return dataTable;
        }


        public static async Task InitializeDb(ReportSysContext context)
        {
            // Укажите путь к вашему Excel-файлу
            string filePath = "C://Users//nikic//Downloads/Данные.xlsx";

            // Загружаем данные из Excel
            DataTable dataTable = LoadExcelFile(filePath);


            //var authusers = new List<AuthUser>
            //{
            //    new AuthUser
            //    {
            //        Login = "test0",
            //        Password = "123",
            //        AccessLevel = 0

            //    },
            //     new AuthUser
            //    {
            //        Login = "test1",
            //        Password = "123",
            //        AccessLevel = 1

            //    },
            //      new AuthUser
            //    {
            //        Login = "test2",
            //        Password = "123",
            //        AccessLevel = 2

            //    },
            //};

            var positions = new List<Position>
            {
                new Position
                {
                    Name = "Ведущий специалист",
                    AccessLevel = 0
                },
                new Position
                {
                    Name = "Начальник отдела",
                    AccessLevel = 1
                },
                new Position
                {
                    Name = "Начальник управления",
                    AccessLevel = 2
                }
            };
            var workschedule = new WorkSchedule
            {
                Arrival = new TimeOnly(8, 30),
                Exit = new TimeOnly(17, 30),
                LunchStart = new TimeOnly(13, 00),
                LunchEnd = new TimeOnly(13, 45)
            };

            var eventTypes = new List<EventType>
            {
                new EventType
                {
                    Name = "Приход"
                },
                new EventType
                {
                    Name = "Уход"
                },
                new EventType
                {
                    Name = "Промежуточная регистрация"
                },

            };

            var unavailabilityTypes = new List<UnavailabilityType>
            {
                new UnavailabilityType
                {
                    Name = "Отпуск"
                },
                new UnavailabilityType
                {
                    Name = "Командировка"
                },
                new UnavailabilityType
                {
                    Name = "Болезнь"
                },
                new UnavailabilityType
                {
                    Name = "Местная командировка"
                },
                new UnavailabilityType
                {
                    Name = "Праздничный день"
                },
            };

            var divisions = new List<Division>
            {
                new Division
                {
                    Name = "Л-Технологии Управление информационной поддержки"
                },
                new Division
                {
                    Name = "Л-Технологии Управление логистических систем"
                },
                new Division
                {
                    Name = "Л-Технологии Управление экономических и финансовых систем"
                },
                new Division
                {
                    Name = "Л-Технологии Управление автоматизацией бухгалтерского учета"
                },
                new Division
                {
                    Name = "Л-Технологии Управление корпоративных платформ и инфроструктура"
                },
            };

            var departments = new List<Department>
            { 
                new Department 
                { 
                    Name = "Л-Технологии Отдел нормативно-справочной информации",
                    Division = divisions[0]
                },
                new Department
                {
                    Name = "Л-Технологии Отдел интеграции",
                    Division = divisions[0]
                },
                new Department
                {
                    Name = "Л-Технологии Отдел оперативной логистики",
                    Division = divisions[1]
                },
                new Department
                {
                    Name = "Л-Технологии Отдел снабжения и сбыта",
                    Division = divisions[1]
                },
                new Department
                {
                    Name = "Л-Технологии Отдел технического обслуживания и ремонта оборудования",
                    Division = divisions[1]
                },
                new Department
                {
                    Name = "Л-Технологии Отдел экономики",
                    Division = divisions[2]
                },
                new Department
                {
                    Name = "Л-Технологии Отдел финансов",
                    Division = divisions[2]
                },
                new Department
                {
                    Name = "Л-Технологии Отдел инвестиционных проектов и договоров",
                    Division = divisions[2]
                },
                new Department
                {
                    Name = "Л-Технологии Отдел бухгалтерского учета",
                    Division = divisions[3]
                },
                new Department
                {
                    Name = "Л-Технологии Отдел учета и отчетности по НДС",
                    Division = divisions[3]
                },
                new Department
                {
                    Name = "Л-Технологии Отдел налогового учёта",
                    Division = divisions[3]
                },
                new Department
                {
                    Name = "Л-Технологии Отдел внеоборотных активов",
                    Division = divisions[3]
                },
                new Department
                {
                    Name = "Л-Технологии Отдел представления и развития систем управленческой отчетности",
                    Division = divisions[4]
                },
            };

            await context.Departments.AddRangeAsync(departments);
            await context.WorkSchedules.AddAsync(workschedule);
            await context.EventTypes.AddRangeAsync(eventTypes);
            await context.UnavailabilityTypes.AddRangeAsync(unavailabilityTypes);

            await context.SaveChangesAsync();
        }
    }
}