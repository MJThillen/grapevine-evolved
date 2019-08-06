package grapevine.model;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlElementWrapper;
import javax.xml.bind.annotation.XmlRootElement;
import java.io.Serializable;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

@XmlRootElement
public class Event implements Serializable {

    public enum EVENT_TYPE {
        none(0),
        action(1),
        plot(2),
        rumor(3),
        cause(4),
        effect(5);


        private int value;

        EVENT_TYPE(int value) {
            this.value = value;
        }

        public int getValue() {
            return value;
        }
    }

    @XmlElement
    private String name;
    @XmlElement
    private LocalDate date;
    @XmlElement
    private String event; //Formerly Outline for Plots
    @XmlElement
    private String subEvent; //Formerly Development for Plot Nodes
    @XmlElement
    private int level;
    @XmlElement
    private EVENT_TYPE eventType;
    @XmlElementWrapper(name="causeList")
    @XmlElement(name="cause")
    private List<Event> causes;
    @XmlElementWrapper(name="effectList")
    @XmlElement(name="effect")
    private List<Event> effects; //Formerly SubRumors for Rumors

    public Event() {
        eventType = EVENT_TYPE.none;
        initialize();
    }

    public Event(final EVENT_TYPE type) {
        eventType = type;
        initialize();
    }

    private void initialize() {
        this.date = LocalDate.now();
        this.level = 0;
        this.causes = new ArrayList<>();
        this.effects = new ArrayList<>();
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getEvent() {
        return event;
    }

    public void setEvent(String event) {
        this.event = event;
    }

    public String getSubEvent() {
        return subEvent;
    }

    public void setSubEvent(String result) {
        this.subEvent = result;
    }

    public List<Event> getCauses() {
        return causes;
    }

    public void setCauses(List<Event> causes) {
        this.causes = causes;
    }

    public List<Event> getEffects() {
        return effects;
    }

    public void setEffects(List<Event> effects) {
        this.effects = effects;
    }

    public EVENT_TYPE getEventType() {
        return eventType;
    }

    public void setEventType(EVENT_TYPE eventType) {
        this.eventType = eventType;
    }

    public LocalDate getDate() {
        return date;
    }

    public void setDate(LocalDate date) {
        this.date = date;
    }

    public void addCause(Event cause) {
        causes.add(cause);
    }

    public void addEffect(Event effect) {
        effects.add(effect);
    }

    public int getLevel() {
        return level;
    }

    public void setLevel(int level) {
        this.level = level;
    }

    public String shortDesc(LocalDate when) {
        if (!when.equals(date)) {
            return date + " " + this.toString();
        } else {
            return this.toString();
        }
    }

    public String toString() {
        String base = "";
        switch(eventType) {
            case none:
                return base;
            case action:
                return name.trim() + " " + event.trim() + " Action";
            case plot:
                return name.trim() + " Plot";
            case rumor:
                if (event.isEmpty()) {
                    return name.trim() + " Rumor";
                } else if (name.endsWith(" Influence")) {
                    return name.substring(0, name.length()-9) + event + " Rumor";
                } else {
                    return name.trim() + " " + event.trim() + " Rumor";
                }
            case cause:
                return "Cause: " + name.trim() + " on " + date;
            case effect:
                return "Effect: " + name.trim() + " on " + date;
        }
        return base;
    }

}
