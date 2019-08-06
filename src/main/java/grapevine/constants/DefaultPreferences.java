package grapevine.constants;

public enum DefaultPreferences {
    //We're going to implement a set of preferences for games, so that they're available throughout the classes,
    //not only in the main Game container.
    // These are key-value pairs, using the Java Preferences API.
    //Given values are the defaults, for when something doesn't have a set preference.

    ENFORCE_HISTORY("EnforceHistory", "true"),
    EXTENDED_HEALTH("ExtendedHealth", "true"),
    FILE_FORMAT("FileFormat", FileFormat.XML.toString()),
    FILE_NAME("FileName", "grapevine-game"),
    LINK_TRAIT_MAXES("LinkTraitMaxes", "true"),
    RANDOM_TRAITS("RandomTraits", "7,5,3,5,5,5,5"),
    ST_COMMENT_END("STCommentEnd", "[/ST]"),
    ST_COMMENT_START("STCommentStart", "[ST]");

    private String name;
    private String value;

    DefaultPreferences(final String name, final String value) {
        this.name = name;
        this.value = value;
    }

    public String getValue() {
        return value;
    }

    public String getName() {
        return name;
    }
}
