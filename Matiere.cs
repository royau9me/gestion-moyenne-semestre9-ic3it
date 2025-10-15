namespace GESTION_MOYENNE_SEMESTRE_9_IC3__IT
{
    public class Matiere
    {
        public string Nom { get; set; } = string.Empty;
        public decimal HeuresDediees { get; set; }
        public int Coefficient { get; set; }
        public decimal Moyenne { get; set; }

        public decimal MoyennePonderee => Moyenne * Coefficient;
    }
}