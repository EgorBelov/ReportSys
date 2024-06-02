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
            //string filePath = "E://HSE_HERNYA//Prac//ReportSys//XlsxFiles//Данные (1).xlsx";

            // Загружаем данные из Excel
            //DataTable dataTable = LoadExcelFile(filePath);


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
                    AccessLevel = 1
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
                    Name = "Вход"
                },
                new EventType
                {
                    Name = "Выход"
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

            var departments = new List<Department>
            {
                new Department
                {
                    Name = "Л-Технологии Управление информационной поддержки"
                },
                new Department
                {
                    Name = "Л-Технологии Управление логистических систем"
                },
                new Department
                {
                    Name = "Л-Технологии Управление экономических и финансовых систем"
                },
                new Department
                {
                    Name = "Л-Технологии Управление автоматизации бухгалтерского учета"
                },
                new Department
                {
                    Name = "Л-Технологии Управление корпоративных платформ и инфроструктура"
                },

                new Department
                {
                    Name = "Л-Технологии Отдел нормативно-справочной информации",
                },
                new Department
                {
                    Name = "Л-Технологии Отдел интеграции",
                },
                new Department
                {
                    Name = "Л-Технологии Отдел оперативной логистики",
                },
                new Department
                {
                    Name = "Л-Технологии Отдел снабжения и сбыта",
                },
                new Department
                {
                    Name = "Л-Технологии Отдел технического обслуживания и ремонта оборудования",
                },
                new Department
                {
                    Name = "Л-Технологии Отдел экономики",
                },
                new Department
                {
                    Name = "Л-Технологии Отдел финансов",
                },
                new Department
                {
                    Name = "Л-Технологии Отдел инвестиционных проектов и договоров",
                },
                new Department
                {
                    Name = "Л-Технологии Отдел бухгалтерского учета",
                },
                new Department
                {
                    Name = "Л-Технологии Отдел учета и отчетности по НДС",
                },
                new Department
                {
                    Name = "Л-Технологии Отдел налогового учета",
                },
                new Department
                {
                    Name = "Л-Технологии Отдел внеоборотных активов",
                },
                new Department
                {
                    Name = "Л-Технологии Отдел представления и развития систем управленческой отчетности",
                }
            };

            var hierarchies = new List<Hierarchy>
            { 
                new Hierarchy 
                { 
                    UpperDepartment = departments[0],
                    LowerDepartment = departments[5],
                },
                new Hierarchy
                {
                    UpperDepartment = departments[0],
                    LowerDepartment = departments[6],
                },
                new Hierarchy
                {
                    UpperDepartment = departments[1],
                    LowerDepartment = departments[7],
                },
                new Hierarchy
                {
                   UpperDepartment = departments[1],
                    LowerDepartment = departments[8],
                },
                new Hierarchy
                {
                    UpperDepartment = departments[1],
                    LowerDepartment = departments[9],
                },
                new Hierarchy
                {
                    UpperDepartment = departments[2],
                    LowerDepartment = departments[10],
                },
                new Hierarchy
                {
                    UpperDepartment = departments[2],
                    LowerDepartment = departments[11],
                },
                new Hierarchy
                {
                    UpperDepartment = departments[2],
                    LowerDepartment = departments[12],
                },
                new Hierarchy
                {
                    UpperDepartment = departments[3],
                    LowerDepartment = departments[13],
                },
                new Hierarchy
                {
                    UpperDepartment = departments[3],
                    LowerDepartment = departments[14],
                },
                new Hierarchy
                {
                    UpperDepartment = departments[3],
                    LowerDepartment = departments[15],
                },
                new Hierarchy
                {
                    UpperDepartment = departments[3],
                    LowerDepartment = departments[16],
                },
                new Hierarchy
                {
                    UpperDepartment = departments[4],
                    LowerDepartment = departments[17],
                },
            };

            await context.Hierarchies.AddRangeAsync(hierarchies);
            await context.WorkSchedules.AddAsync(workschedule);
            await context.EventTypes.AddRangeAsync(eventTypes);
            await context.UnavailabilityTypes.AddRangeAsync(unavailabilityTypes);
            await context.Positions.AddRangeAsync(positions);

            await context.SaveChangesAsync();
        }
    }
}