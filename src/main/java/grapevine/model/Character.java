package grapevine.model;

import grapevine.constants.QueryConstants.*;

public interface Character {
    String status = "";
    String name = "";
    String race = "";
    String group = "";
    String subGroup = "";

    boolean isActive();
    String getStatus();
    String getName();
    String getRace();
    String getGroup();
    String getSubGroup();
    void setStatus(final String status);
    void setName(final String name);
    void setRace(final String race);
    void setGroup(final String group);
    void setSubGroup(final String group);
    TraitList<Trait> getValue(QueryKeys key);

}
