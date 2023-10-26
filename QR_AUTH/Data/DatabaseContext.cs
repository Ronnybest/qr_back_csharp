using Microsoft.EntityFrameworkCore;
using QR_AUTH.Models;

namespace QR_AUTH.Data;

public class DatabaseContext : DbContext
{
    public DatabaseContext(DbContextOptions options) : base(options: options)
    {
    }

    public virtual DbSet<AuthModel> Auths { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        optionsBuilder.UseNpgsql();
    }
}