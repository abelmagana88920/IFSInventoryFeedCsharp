using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;

namespace InventoryFeedService
{
    public partial class IFSReportingContext : DbContext
    {
        public IFSReportingContext()
            : base("name=IFSReportingContext")
        {
        }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }

        public DbSet<tblInvoiceLinesMaster> tblInvoiceLinesMasters { get; set; }
        public DbSet<tblInventoryFeed> tblInventoryFeeds { get; set; }
        public DbSet<tblInventoryLog> tblInventoryLogs { get; set; }
        public DbSet<tblInventoryFeedProcess> tblInventoryFeedProcesses { get; set; }
    }
}
