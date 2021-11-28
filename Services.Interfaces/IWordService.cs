using JsonToWord.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace JsonToWord.Services.Interfaces
{
    public interface IWordService
    {
        string Create(WordModel _wordModel);
    }
}
