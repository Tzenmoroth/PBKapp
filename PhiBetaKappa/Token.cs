using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhiBettaKappa
{
    public enum TokenType
    {
        IDtag,
        NameTag,
        CollegeTag,
        TermTag,
        CumTag,
        IDnum,
        ProgressTag,
        OverallTag,
        TransferTag,
        InstitutionTag,
        TransTotalTag,
        InsTotalTag,
        Value,
        CourseSub, // Course SUBJECT
        CourseNum, // Course NUMBER
        CourseTitle,
        Grade,
        Cred,
        Null // Unknown token
    };

    public class Token
    {
        public String data;

        public TokenType type;

        public Token(String iData, TokenType iType)
        {
            data = iData;
            type = iType;
        }

        public static List<int> indexesOfType(List<Token> list, TokenType typeToFind){
            List<int> output = new List<int>();
            for (int i = 0; i < list.Count; i++) if (list[i].type == typeToFind) output.Add(i);
            return output;
        }
    }
}
