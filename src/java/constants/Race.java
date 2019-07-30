package java.constants;

public enum Race {

    DELETE("", -1),
    NONE("none", 0),
    ALL("all", 1),
    VAMPIRE("Vampire", 2),
    WEREWOLF("Werewolf", 3),
    MORTAL("Mortal", 4),
    CHANGELING("Changeling", 5),
    WRAITH("Wraith", 6),
    MAGE("Mage", 7),
    FERA("Fera", 8),
    VARIOUS("Various", 9),
    MUMMY("Mummy", 10),
    KUEIJIN("KueiJin", 11),
    HUNTER("Hunter", 12),
    DEMON("Demon", 13);

    private int number;
    private String name;

    private Race(String name,
                 int number) {
        this.number = number;
        this.name = name;
    }

    public int getNumber() {
        return number;
    }

    public String getName() {
        return name;
    }

    @Override
    public String toString() {
        return name;
    }
}
