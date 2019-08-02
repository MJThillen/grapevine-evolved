package grapevine;

import grapevine.constants.ListDisplay;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import java.util.LinkedList;

@XmlRootElement
public class TraitList<E> extends LinkedList<E> {

    @XmlElement
    private String name;
    @XmlElement
    private int traitTotal;
    @XmlElement
    private ListDisplay displayType;
    @XmlElement
    private boolean alphabetized;
    @XmlElement
    private boolean atomic;
    @XmlElement
    private boolean negative;

    public TraitList(final String name,
                     final boolean alphabetized,
                     final boolean negative,
                     final boolean atomic,
                     final ListDisplay displayType) {
        super();
        this.name = name;
        this.alphabetized = alphabetized;
        this.negative = negative;
        this.atomic = atomic;
        this.displayType = displayType;
    }

    public TraitList() {
        super();
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getTraitTotal() {
        return traitTotal;
    }

    public void setTraitTotal(int traitTotal) {
        this.traitTotal = traitTotal;
    }

    public ListDisplay getDisplayType() {
        return displayType;
    }

    public void setDisplayType(ListDisplay displayType) {
        this.displayType = displayType;
    }

    public boolean isAlphabetized() {
        return alphabetized;
    }

    public void setAlphabetized(boolean alphabetized) {
        this.alphabetized = alphabetized;
    }

    public boolean isAtomic() {
        return atomic;
    }

    public void setAtomic(boolean atomic) {
        this.atomic = atomic;
    }

    public boolean isNegative() {
        return negative;
    }

    public void setNegative(boolean negative) {
        this.negative = negative;
    }

    public int count() {
        return atomic? this.size() : traitTotal;
    }


    //ToDo: Add binary input/output methods.
}
