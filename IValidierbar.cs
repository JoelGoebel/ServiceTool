public interface IValidierbar
{

    //Gibt true zurück, wenn Pflichtfelder fehlen.
    // Gibt zusätzlich eine Fehlermeldung zurück.

    bool HatFehlendePflichtfelder(out string fehlermeldung);
}
