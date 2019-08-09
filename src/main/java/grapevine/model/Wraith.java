package grapevine.model;

public class Wraith extends Character {



    /**
     * A wrapper method for readability.
     * @return The group value where we stored the ethnos.
     * */
    public String getEthnos() {
        return this.getGroup();
    }

    /**
     * A wrapper method for readability
     * @param ethnos The value of ethnos to store in the group variable.
     */
    public void setEthnos(String ethnos) {
        this.setGroup(ethnos);
    }

    /**
     * A wrapper method for readability.
     * @return the subGroup value where we stored the guild.
     */
    public String getGuild() {
        return this.getSubGroup();
    }

    /**
     * A wrapper method for readability
     * @param guild the value of guild to store in the subgroup variable.
     */
    public void setGuild(String guild) {
        this.setSubGroup(guild);
    }
}
