using System.Collections.Generic;

namespace JsonToWord.Models
{
    public class WordTableRow
    {
        public List<WordTableCell> Cells { get; set; }
        public bool MergeToOneCell { get; set; }
        public int NumberOfCellsToMerge { get; set; }
    }
}