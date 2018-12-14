﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ppk5_v2
{
    class Tests
    {
        public static string[] testArray500()
        {
            string[] array = new string[] { "150111:1409", "50:21:0150202:210", "50:21:0150206:174",
                "50:21:0150301:398", "50:21:0150301:403", "50:21:0150302:265", "50:21:0150302:333",
                "50:21:0150304:402", "50:26:0000000:4657", "50:26:0000000:51415", "50:26:0000000:6642",
                "50:26:0000000:8144", "50:26:0000000:8145", "50:26:0000000:8262", "50:26:0000000:9155",
                "50:26:0000000:9164", "50:26:0040903:452", "50:26:0050702:136", "50:26:0100103:855",
                "50:26:0100204:258", "50:26:0110703:201", "50:26:0110704:710", "50:26:0110727:185",
                "50:26:0110727:194", "50:26:0130201:1074", "50:26:0130201:1084", "50:26:0130201:1086",
                "50:26:0130202:906", "50:26:0140401:1070", "50:26:0140401:1557", "50:26:0140401:803",
                "50:26:0140402:514", "50:26:0140410:387", "50:26:0140410:415", "50:26:0140427:503",
                "50:26:0140502:580", "50:26:0150101:108", "50:26:0150202:254", "50:26:0150301:496",
                "50:26:0150601:354", "50:26:0151302:1356", "50:26:0151305:391", "50:26:0151802:117",
                "50:26:0152001:125", "50:26:0152003:247", "50:26:0170101:244", "50:26:0170101:253",
                "50:26:0170102:575", "50:26:0170103:555", "50:26:0170103:615", "50:26:0170103:647",
                "50:26:0170103:677", "50:26:0170103:686", "50:26:0170103:704", "50:26:0170103:708",
                "50:26:0170103:735", "50:26:0170103:749", "50:26:0170103:782", "50:26:0170103:814",
                "50:26:0170103:843", "50:26:0170103:845", "50:26:0170103:856", "50:26:0170103:908",
                "50:26:0170103:996", "50:26:0170104:734", "50:26:0170104:735", "50:26:0170105:7",
                "50:26:0170201:502", "50:26:0170202:690", "50:26:0170205:131", "50:26:0170401:617",
                "50:26:0170402:1206", "50:26:0170402:1883", "50:26:0170402:1895", "50:26:0170402:2234",
                "50:26:0170504:702", "50:26:0170504:745", "50:26:0170508:556", "50:26:0170519:195",
                "50:26:0170704:322", "50:26:0170704:323", "50:26:0170704:393", "50:26:0170704:396",
                "50:26:0170704:770", "50:26:0170704:862", "50:26:0171110:236", "50:26:0180501:293",
                "50:26:0180501:349", "50:26:0180501:470", "50:26:0180503:397", "50:26:0180505:374",
                "50:26:0180511:218", "50:26:0190201:346", "50:26:0190301:566", "50:26:0190401:610",
                "50:26:0190402:287", "50:26:0190402:314", "50:26:0190802:418", "50:26:0190805:192",
                "50:26:0190901:251", "50:26:0190901:424", "50:26:0190901:796", "50:26:0190901:801",
                "50:26:0190905:434", "50:26:0190905:528", "50:26:0190905:562", "50:26:0190905:579",
                "50:26:0190905:651", "50:26:0190905:840", "50:26:0190905:979", "50:26:0190907:128",
                "50:26:0190907:159", "50:26:0190907:186", "50:26:0190907:327", "50:26:0190915:122",
                "50:26:0191002:234", "50:26:0191201:774", "50:26:0191201:945", "50:26:0191204:267",
                "50:26:0191207:415", "50:26:0191215:380", "50:26:0191233:16", "50:26:0191401:1052",
                "50:26:0191401:1054", "50:26:0191401:1055", "50:26:0191401:1124", "50:26:0191401:711",
                "50:26:0191403:516", "50:27:0000000:130059", "50:27:0000000:130518", "50:27:0000000:15991",
                "50:27:0000000:16024", "50:27:0000000:18017", "50:27:0000000:18018", "50:27:0000000:19382",
                "50:27:0000000:19383", "50:27:0000000:19385", "50:27:0000000:19386", "50:27:0000000:19444",
                "50:27:0000000:19445", "50:27:0000000:23647", "50:27:0000000:24532", "50:27:0000000:2828",
                "50:27:0000000:28622", "50:27:0000000:3106", "50:27:0000000:31924", "50:27:0000000:32087",
                "50:27:0000000:32582", "50:27:0000000:35113", "50:27:0000000:35378", "50:27:0000000:37048",
                "50:27:0000000:38894", "50:27:0000000:3896", "50:27:0000000:39736", "50:27:0000000:40109",
                "50:27:0000000:4152", "50:27:0000000:4400", "50:27:0000000:45415", "50:27:0000000:47171",
                "50:27:0000000:51330", "50:27:0000000:51564", "50:27:0000000:53523", "50:27:0000000:53581",
                "50:27:0000000:53650", "50:27:0000000:56348", "50:27:0000000:56440", "50:27:0000000:56762",
                "50:27:0000000:6396", "50:27:0000000:6397", "50:27:0000000:6398", "50:27:0020103:174",
                "50:27:0020114:456", "50:27:0020115:658", "50:27:0020121:207", "50:27:0020201:159",
                "50:27:0020202:582", "50:27:0020219:137", "50:27:0020322:311", "50:27:0020401:276",
                "50:27:0020410:670", "50:27:0020411:431", "50:27:0020415:402", "50:27:0020423:335",
                "50:27:0020424:339", "50:27:0020437:279", "50:27:0020463:393", "50:27:0020471:888",
                "50:27:0030125:549", "50:27:0030146:156", "50:27:0030148:65", "50:27:0030201:1088",
                "50:27:0030201:1387", "50:27:0030220:152", "50:27:0030228:76", "50:27:0030321:281",
                "50:27:0030405:271", "50:27:0030405:456", "50:27:0030426:1607", "50:27:0030514:274",
                "50:27:0030522:488", "50:27:0030603:230", "50:27:0030618:156", "50:27:0030635:468",
                "50:27:0030643:174", "50:27:0040104:450", "50:27:0040215:426", "50:27:0040301:231",
                "50:42:0000000:76358", "50:42:0000000:76359", "50:42:0000000:78567", "50:54:0010201:289",
                "50:54:0010201:361", "50:54:0010201:578", "50:54:0010201:723", "50:54:0010201:737",
                "50:54:0010204:9", "50:54:0020103:29", "50:54:0020103:71", "50:54:0020103:72",
                "50:54:0020105:247", "50:54:0020306:18", "50:54:0020312:41", "50:54:0020314:37",
                "50:54:0020404:25", "50:54:0020409:183", "50:54:0020409:214", "50:61:0000000:754",
                "50:61:0000000:949", "50:61:0010106:60", "50:61:0010118:16", "50:61:0010121:116",
                "50:61:0010121:164", "50:61:0010121:89", "50:61:0010122:138", "50:61:0010122:453",
                "50:61:0010122:455", "50:61:0010122:97", "50:61:0010123:66", "50:61:0010128:99",
                "50:61:0010201:667", "50:61:0010201:709", "50:61:0020101:242", "50:61:0020101:250",
                "50:61:0020203:55", "50:61:0020211:167", "50:61:0020222:187", "50:61:0020223:74",
                "50:61:0020225:103", "50:61:0020227:56", "50:61:0020239:91", "50:61:0020245:70",
                "50:61:0020250:34", "50:61:0020268:66", "50:61:0020277:55", "50:61:0020278:49",
                "50:61:0020278:60", "50:61:0030106:46", "77:00:0000000:16199", "77:00:0000000:16201",
                "77:00:0000000:16204", "77:00:0000000:16211", "77:00:0000000:16212", "77:00:0000000:16215",
                "77:00:0000000:16216", "77:00:0000000:16217", "77:00:0000000:16218", "77:00:0000000:16219",
                "77:00:0000000:16221", "77:00:0000000:16222", "77:00:0000000:16223", "77:00:0000000:16228",
                "77:00:0000000:16230", "77:00:0000000:16234", "77:00:0000000:16237", "77:00:0000000:16238",
                "77:00:0000000:16240", "77:00:0000000:16242", "77:00:0000000:16244", "77:00:0000000:16246",
                "77:00:0000000:16251", "77:00:0000000:16255", "77:00:0000000:16258", "77:00:0000000:16260",
                "77:00:0000000:16261", "77:00:0000000:16263", "77:00:0000000:16264", "77:00:0000000:16265",
                "77:00:0000000:16266", "77:00:0000000:16267", "77:00:0000000:16269", "77:00:0000000:16270",
                "77:00:0000000:16271", "77:00:0000000:16282", "77:00:0000000:16285", "77:00:0000000:16286",
                "77:00:0000000:16287", "77:00:0000000:16288", "77:00:0000000:16289", "77:00:0000000:16290",
                "77:00:0000000:16291", "77:00:0000000:16292", "77:00:0000000:16293", "77:00:0000000:16294",
                "77:00:0000000:16295", "77:00:0000000:16296", "77:00:0000000:16297", "77:00:0000000:16298",
                "77:00:0000000:16299", "77:00:0000000:16301", "77:00:0000000:16303", "77:00:0000000:16305",
                "77:00:0000000:16306", "77:00:0000000:16307", "77:00:0000000:16308", "77:00:0000000:16309",
                "77:00:0000000:16312", "77:00:0000000:16313", "77:00:0000000:16314", "77:00:0000000:16315",
                "77:00:0000000:16316", "77:00:0000000:16317", "77:00:0000000:16319", "77:00:0000000:16320",
                "77:00:0000000:16321", "77:00:0000000:16325", "77:00:0000000:16352", "77:00:0000000:16353",
                "77:00:0000000:16356", "77:00:0000000:16357", "77:00:0000000:16358", "77:00:0000000:16359",
                "77:00:0000000:16360", "77:00:0000000:16361", "77:00:0000000:16362", "77:00:0000000:16364",
                "77:00:0000000:16365", "77:00:0000000:16366", "77:00:0000000:16367", "77:00:0000000:16370"}; 
            return array;
        }
        /// <summary>
        /// Cast to List<T> before using it!
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<Elem> testElem500()
        {
            var temp = testArray500();
            
            foreach (var val in temp)
            {
                yield return new Elem(val);
            }
        }
    }
}
