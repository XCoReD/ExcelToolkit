using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Tools
{
    //advice from https://stackoverflow.com/questions/4839470/array-initialization-with-default-constructor
    public static class ArrayInitializer
    {
        public static T[] Populate<T>(this T[] array, Func<T> provider)
        {
            for (int i = 0; i < array.Length; i++)
            {
                array[i] = provider();
            }
            return array;
        }
    };
}
