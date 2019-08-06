package grapevine.constants;

public class QueryConstants {

    private QueryConstants() {
        // An empty private constructor restricts instantiation.
    }

    public enum QueryKeyType {
        ERROR(-1), //No Type: Error or Uninitialized
        FIELD(0), //String
        NUMBER(1), //Integer or Double
        TRAIT_LIST(2), //TraitList
        DATE(3), //LocalDate
        BOOLEAN(4); //Boolean

        private int value;

        QueryKeyType(int value) {
            this.value = value;
        }

        public int getValue() {
            return value;
        }
    }

    public enum Statistic { //An enumeration of the statistics that can be run
        DISTRIBUTION(0), // Distribution (General)
        DISTINCT(1), // Distinct Trait Distribution
        SPECIFIC(2), // Specific Trait Distribution
        MAXIMA(3), //Maxima of trait lists or number data
        SUMS(4); //Sums of trait lists or number data

        private int value;

        Statistic(int value) {
            this.value = value;
        }

        public int getValue() {
            return value;
        }
    }

    public enum QueryCompare {
        CONTAINS(0),
        EQUALS(1),
        GREATER_EQUAL(2),
        GREATER(3),
        LESS(4),
        LESS_EQUAL(5),
        CONTAINS_EXACTLY(6),
        CONTAINS_GREATER_EQUAL(7),
        CONTAINS_GREATER(8),
        CONTAINS_LESS(9),
        CONTAINS_LESS_EQUAL(10),
        TOTALS(11),
        TOTALS_GREATER_EQUAL(12),
        TOTALS_GREATER(13),
        TOTALS_LESS_EQUAL(14),
        TOTALS_LESS(15),
        CONTAINS_NOTE(16),
        IS_TRUE(17),
        IS_FALSE(18);

        private int value;

        QueryCompare(int value) {
            this.value = value;
        }

        public int getValue() {
            return value;
        }
    }

    public enum QueryInventory {
        NONE(0),
        CHARACTERS(1),
        PLAYERS(2),
        ITEMS(4),
        LOCATIONS(8),
        ACTIONS(16),
        PLOTS(32),
        RUMORS(64),
        ROTES(128);

        private int value;

        QueryInventory(int value) {
            this.value = value;
        }

        public int getValue() {
            return value;
        }
    }

    public enum QueryKeys {
        ABILITIES("abilities"),
        ACCESS("access"),
        ACTION("action"),
        ACTIVE("active"),
        ADDRESS("address"),
        AFFINITY("affinity"),
        AMENTI("amenti"),
        ANGST("angst"),
        APPEARANCE("appearance"),
        ARCANOI("arcanoi"),
        ARETE("arete"),
        ARTS("arts"),
        ASSOCIATION("association"),
        AURA("aura"),
        AURA_BONUS("aurabonus"),
        AUSPICE("auspice"),
        AVAILABILITY("availability"),
        BA("ba"),
        BACKGROUND("background"),
        BANALITY("banality"),
        BIOGRAPHY("biography"),
        BLOOD("blood"),
        BONDS("bonds"),
        BONUS("bonus"),
        BOONS("boons"),
        BREED("breed"),
        BROOD("brood"),
        CABAL("cabal"),
        CAMP("camp"),
        CHANGE("change"),
        CHANGE_TEXT("changetext"),
        CLAN("clan"),
        CLASS("class"),
        CONCEALABILITY("conceal"),
        CONSCIENCE("conscience"),
        CONVICTION("conviction"),
        CORPUS("corpus"),
        COTERIE("coterie"),
        COUNT("count"),
        COURAGE("courage"),
        COURT("court"),
        CREED("creed"),
        DAMAGE_AMOUNT("damageamount"),
        DAMAGE_TYPE("damagetype"),
        DARK_PASSIONS("darkpassions"),
        DATE("date"),
        DEATH("death"),
        DEMEANOR("demeanor"),
        DEMON_CHI("demonchi"),
        DERANGEMENTS("derangements"),
        DESCRIPTION("description"),
        DEV_DATE("devdate"),
        DEVELOPMENT("development"),
        DHARMA("dharma"),
        DIRECTION("direction"),
        DISCIPLINES("disciplines"),
        DURATION("duration"),
        EARNED("earned"),
        EDGES("edges"),
        EMAIL("email"),
        END_DATE("enddate"),
        EQUIPMENT("equipment"),
        ESSENCE("essence"),
        ETHNOS("ethnos"),
        FACTION("faction"),
        FAITH("faith"),
        FEATURES("features"),
        FERA("fera"),
        FETTERS("fetters"),
        FLAWS("flaws"),
        FOCI("foci"),
        GAUNTLET("gauntlet"),
        GENERATION("generation"),
        GIFTS("gifts"),
        GLAMOUR("glamour"),
        GLORY("glory"),
        GLORY_TRAITS("glorytraits"),
        GNOSIS("gnosis"),
        GRADES("grades"),
        GROUP("group"),
        GROWTH("growth"),
        GUANXI("guanxi"),
        GUILD("guild"),
        HANDLE("handle"),
        HAUNT("haunt"),
        HEALTH("health"),
        HEKAU("hekau"),
        HISTORY("history"),
        HONOR("honor"),
        HONOR_TRAITS("honortraits"),
        HOUSE("house"),
        HUMANITY("humanity"),
        HUN("hun"),
        ID("id"),
        INFLUENCES("influences"),
        INHERITANCE("inheritance"),
        INTEGRITY("integrity"),
        JOY("joy"),
        KA("ka"),
        KEY_CHARACTERS("keycharacters"),
        KITH("kith"),
        KJ_BALANCE("kjbalance"),
        LAST_MODIFIED("lastmodified"),
        LEGION("legion"),
        LEVEL("level"),
        LIFE("life"),
        LINKS("links"),
        LOCATIONS("loctions"),
        LORES("lores"),
        M_BALANCE("mbalance"),
        MEMORY("memory"),
        MENTAL("mental"),
        MENTAL_MAX("mentalmax"),
        MENTAL_NEG("mentalneg"),
        MERCY("mercy"),
        MERITS("merits"),
        MISCELLANEOUS("miscellaneous"),
        MOTIVATION("motivation"),
        NAME("name"),
        NARRATOR("narrator"),
        NATURE("nature"),
        NEGATIVES("negatives"),
        NEXT_GAME("nextgame"),
        NEXT_NOTES("nextnotes"),
        NEXT_PLACE("nextplace"),
        NEXT_TIME("nexttime"),
        NPC("npc"),
        NOTE("note"),
        NOTES("notes"),
        NOTORIETY("notoriety"),
        OATHS("oaths"),
        OTHER("other"),
        OUTLINE("outline"),
        OWED("owed"),
        OWNER("owner"),
        PACK("pack"),
        PARADOX("paradox"),
        PARTNER("partner"),
        PASSIONS("passions"),
        PATH("path"),
        PATH_TRAITS("pathtraits"),
        PATHOS("pathos"),
        PHONE("phone"),
        PHYSICAL("physical"),
        PHYSICAL_MAX("physicalmax"),
        PHYSICAL_NEG("physicalneg"),
        PLANE("plane"),
        PLAY_STATUS("playstatus"),
        PLAYER("player"),
        PO("po"),
        PO_ARCHETYPE("poarchetype"),
        POSITION("position"),
        POWERS("powers"),
        PP_EARNED("ppearned"),
        PP_UNSPENT("ppunspent"),
        QUINTESSENCE("quintessence"),
        RACE("race"),
        RAGE("rage"),
        RANDOM("random"),
        RANK("rank"),
        REALMS("realms"),
        REASON("reason"),
        REGNANT("regnant"),
        REGRET("regret"),
        REPUTATION("reputation"),
        RESONANCE("resonance"),
        RESULT("result"),
        RITES("rites"),
        RITUALS("rituals"),
        ROTES("rotes"),
        RUMOR("rumor"),
        SECT("sect"),
        SECURITY("security"),
        SECURITY_RETESTS("securityretests"),
        SECURITY_TRAITS("securitytraits"),
        SEELIE_LEGACY("seelielegacy"),
        SEEMING("seeming"),
        SEKHEM("sekhem"),
        SELF_CONTROL("selfcontrol"),
        SHADOW_ARCHETYPE("shadowarchetype"),
        SHADOW_PLAYER("shadowplayer"),
        SIRE("sire"),
        SOCIAL("social"),
        SOCIAL_MAX("socialmax"),
        SOCIAL_NEG("socialneg"),
        SPELLS("spells"),
        SPHERES("spheres"),
        START_DATE("startdate"),
        STATION("station"),
        STATUS("status"),
        SUB_CLASS("subclass"),
        SUB_TYPE("subtype"),
        SUBGROUP("subgroup"),
        TEMP_ANGST("tempangst"),
        TEMP_ARETE("temparete"),
        TEMP_BA("tempba"),
        TEMP_BANALITY("tempbanality"),
        TEMP_BLOOD("tempblood"),
        TEMP_CONSCIENCE("tempconscience"),
        TEMP_CONVICTION("tempconviction"),
        TEMP_CORPUS("tempcorpus"),
        TEMP_COURAGE("tempcourage"),
        TEMP_DEMON_CHI("tempdemonchi"),
        TEMP_FAITH("tempfaith"),
        TEMP_GLAMOUR("tempglamour"),
        TEMP_GLORY("tempglory"),
        TEMP_GNOSIS("tempgnosis"),
        TEMP_HONOR("temphonor"),
        TEMP_HUN("temphun"),
        TEMP_INTEGRITY("tempintegrity"),
        TEMP_JOY("tempjoy"),
        TEMP_KA("tempka"),
        TEMP_M_BALANCE("tempmbalance"),
        TEMP_MEMORY("tempmemory"),
        TEMP_MERCY("tempmercy"),
        TEMP_PARADOX("tempparadox"),
        TEMP_PATH_TRAITS("temppathtraits"),
        TEMP_PATHOS("temppathos"),
        TEMP_PO("temppo"),
        TEMP_QUINTESSENCE("tempquintessence"),
        TEMP_RAGE("temprage"),
        TEMP_SEKHEM("tempsekhem"),
        TEMP_SELF_CONTROL("tempselfcontrol"),
        TEMP_TORMENT("temptorment"),
        TEMP_TRUE_FAITH("temptruefaith"),
        TEMP_VISION("tempvision"),
        TEMP_WILLPOWER("tempwillpower"),
        TEMP_WISDOM("tempwisdom"),
        TEMP_YANG("tempyang"),
        TEMP_YIN("tempyin"),
        TEMP_ZEAL("tempzeal"),
        TEMPERS("tempers"),
        THORNS("thorns"),
        THRESHOLD("threshold"),
        TITLE("title"),
        TORMENT("torment"),
        TOTAL("total"),
        TOTEM("totem"),
        TRADITION("tradition"),
        TRIBE("tribe"),
        TRUE_FAITH("truefaith"),
        TYPE("type"),
        UMBRA("umbra"),
        UNSEELIE_LEGACY("unseelielegacy"),
        UNSPENT("unspent"),
        UNUSED("unused"),
        USUAL_PLACE("usualplace"),
        USUAL_TIME("usualtime"),
        VALUE("value"),
        VISAGE("apocalypticform"),
        VISION("vision"),
        WEB_SITE("website"),
        WHERE("where"),
        WILLPOWER("willpower"),
        WISDOM("wisdom"),
        WISDOM_TRAITS("wisdomtraits"),
        XP_EARNED("xpearned"),
        XP_UNSPENT("xpunspent"),
        YANG("yang"),
        YIN("yin"),
        ZEAL("zeal"),
    
        DEFAULT("");
        
        private String value;

        QueryKeys(String value) {
            this.value = value;
        }

        public String getValue() {
            return value;
        }
    }
}
