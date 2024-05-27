﻿using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using ReportSys.DAL.Entities;

namespace ReportSys.DAL
{
    public class ReportSysContext : DbContext
    {

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Employee>()
                .HasOne(e => e.AuthUser)
                .WithOne(au => au.Employee)
                .HasForeignKey<Employee>(e => e.AuthUserId);

            // Или, если AuthUser должен содержать внешний ключ:
            // modelBuilder.Entity<AuthUser>()
            //     .HasOne(au => au.Employee)
            //     .WithOne(e => e.AuthUser)
            //     .HasForeignKey<AuthUser>(au => au.Id);
        }

        public DbSet<AuthUser> AuthUsers => Set<AuthUser>();
        public DbSet<Employee> Employees => Set<Employee>();
        public DbSet<Department> Departments => Set<Department>();
        public DbSet<Position> Positions => Set<Position>();
        public DbSet<Division> Divisions => Set<Division>();
        public DbSet<WorkSchedule> WorkSchedules => Set<WorkSchedule>();
        public DbSet<EventType> EventTypes => Set<EventType>();
        public DbSet<Event> Events => Set<Event>();
        public DbSet<UnavailabilityType> UnavailabilityTypes => Set<UnavailabilityType>();
        public DbSet<Unavailability> Unavailabilitys => Set<Unavailability>();

        public ReportSysContext(DbContextOptions<ReportSysContext> options)
        : base(options)
        {

        }



    }
}
