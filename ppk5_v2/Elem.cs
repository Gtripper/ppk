

namespace ppk5_v2
{
    /// <summary>
    /// 
    /// </summary>
    public class Elem
    {
        public string cad_num;
        public OKS oks;

        public Elem(string cad_num)
        {
            this.cad_num = cad_num;
        }

        public Elem(string cad_num, OKS oks)
        {
            this.oks = oks;
            this.cad_num = cad_num;
        }
    }
}
