namespace Model
{
    public class Director
    {
        public int Id { get; set; }
        public string FullName { get; set; } = string.Empty;

        public override string ToString() => FullName;
    }
}
