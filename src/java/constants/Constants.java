package java.constants;

public final class Constants {

    private Constants() {
        // An empty private constructor restricts instantiation.
    }

    //Title of current version of Grapevine
    public static final String GRAPEVINE_CAPTION = "Grapevine Evolved";
    //Used to identify file types and versions
    public static final Double CURRENT_VERSION = 4.0;

    //Recent Search Text
    public static final String RECENT_SEARCH_NAME = "Most Recent Search";
    //Backup File Name
    //TODO: Figure out file extension plan
    public static final String BACKUP_FILE_NAME = "~Autosave.gv4";

    //Names of Special Email Recipients
    public static final String SPECIAL_EMAIL = "[Send_to_Select]";

    //Character Values Keys - see PublicQueryKeys module
    public static final String XE_KEY = "Gv3.0XK!";

    //Important Strings dealing with actions and rumors
    public static final String BASIC_SUBACTION_NAME = "Personal";
    public static final String PUBLIC_RUMOR_TITLE = "Public Knowledge";

    //Binary File Header Data, used to identify binary file types and versions
    public static final int BINARY_HEADER_LENGTH = 4;
    public static final String BINARY_HEADER_GAME = "GVBG";
    public static final String BINARY_HEADER_MENU = "GVBM";
    public static final String BINARY_HEADER_EXCHANGE = "GVBE";

    //Locations of Grapevine web resources
    //TODO: Consider what/where Grapevine Evolved is going to go.
    public static final String URL_MAIN = "http://www.GrapevineLARP.com/";
    public static final String URL_HELP = "http://www.GrapevineLARP.com/help.shtml";

    //Filenames for default files
    public static final String DEFAULT_MENU_FILE = "Grapevine Menus.gvm";
    public static final String DEFAULT_ITEMS_FILE = "New Game Items.gex";

    //Game File Version Headers for each file format.
    public static final String GAME_FILE_VERSION_0 = "<-Grapevine II Game File->";
    public static final String GAME_FILE_2_0 = "<-Grapevine 2.0 Game File / Format 1->";
    public static final String GAME_FILE_2_0V2 = "<-Grapevine 2.0 Game File / Format 2->";
    public static final String GAME_FILE_2_1 = "<-Grapevine 2.1 Game File / Format 1->";
    public static final String GAME_FILE_2_2 = "<-Grapevine 2.2 Game File / Format 1->";
    public static final String GAME_FILE_2_3 = "<-Grapevine 2.3 Game File / Format 1->";

    public static final String EXCHANGE_FILE_2_2 = "<-Exchange File / Grapevine 2.2 / Format 1 ->";
    public static final String EXCHANGE_FILE_2_3 = "<-Exchange File / Grapevine 2.3 / Format 1 ->";

    //The item in the status menus that means "Active"
    public static final String ACTIVE_STATUS = "Active";

    //Point Type Tags - A value that controls tracking player points or XP
    public static final String PM_EXPERIENCE = "E";
    public static final String PM_PLAYER_POINTS = "P";

    //Health Level Constants
    public static final String HEALTH_LEVEL_0 = "Healthy";
    public static final String HEALTH_LEVEL_1 = "Bruised";
    public static final String HEALTH_LEVEL_2 = "Wounded";
    public static final String HEALTH_LEVEL_3 = "Incapacitated";
    public static final String HEALTH_LEVEL_4 = "Mortally Wounded";
    public static final int MIN_HEALTH = 0;
    public static final int MAX_HEALTH = 4;

    //Standard Template Names
    public static final String TN_CHARACTER_SHEET_SUFFIX = " Character Sheet";
    public static final String TN_ACTION_RUMOR = "Action and Rumor Report";
    public static final String TN_MASTER_ACTION = "Master Action Report";
    public static final String TN_MASTER_RUMOR = "Master Rumor Report";
    public static final String TN_PLOT = "Plot Report";
    public static final String TN_CHARACTER_SHEETS = "Character Sheets";
    public static final String TN_CHARACTER_ROSTER = "Character Roster";
    public static final String TN_EQUIPMENT = "Character Equipment";
    public static final String TN_SIGN_IN = "Sign-In Sheet";
    public static final String TN_ITEM_CARDS = "Item Cards";
    public static final String TN_ROTE_CARDS = "Rote Cards";
    public static final String TN_LOCATION_CARDS = "Location Cards";
    public static final String TN_XP_HISTORY = "Experience History";
    public static final String TN_PP_HISTORY = "Player Point History";
    public static final String TN_PLAYER_ROSTER = "Player Roster";
    public static final String TN_GAME_CALENDAR = "Game Calendar";
    public static final String TN_SEARCH = "Search Report";
    public static final String TN_STATISTICS = "Statistics Report";
    public static final String TN_MERITS_FLAWS = "Merits and Flaws Report";
    public static final String TN_INFLUENCE = "Influence Report";

    public enum OUTPUT_ID_CONSTANTS {
        none(-1),
        traitList(0),
        history(1),
        plot(2),
        action(3),
        rumor(4),
        calendar(5);

        private int value;

        private OUTPUT_ID_CONSTANTS(int value) {
            this.value = value;
        }

        public int getValue() {
            return value;
        }
    }

    public enum OUTPUT_SELECTION_CONSTANTS {
        min(1),
        players(1),
        characters(2),
        items(3),
        rotes(4),
        locations(5),
        actions(6),
        plots(7),
        rumors(8),
        search(9),
        statistics(10),
        max(8);

        private int value;

        private OUTPUT_SELECTION_CONSTANTS(int value) {
            this.value = value;
        }

        public int getValue() {
            return value;
        }
    }
}
