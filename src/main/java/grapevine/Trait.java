package grapevine;

public class Trait implements Comparable{
    private String name;
    private int total;
    private String note;

    public Trait(){
        name = "";
        total = 0;
        note = "";
    }
    public Trait(final String name, final int total, final String note) {
        this.name = name;
        this.total = total;
        this.note = note;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getTotal() {
        return total;
    }

    public void setTotal(int total) {
        this.total = total;
    }

    public String getNote() {
        return note;
    }

    public void setNote(String note) {
        this.note = note;
    }

    @Override
    public int compareTo(Object o) {
        Trait compTrait = (o instanceof Trait) ? ((Trait) o) : null;
        if (compTrait != null) {
            return this.name.compareTo(compTrait.getName());
        } else {
            return -1;
        }
    }
}
