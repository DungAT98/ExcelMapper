using System.Collections.Generic;

namespace ExcelReaderMapper.Extensions
{
    public static class EnumerableExtensions
    {
        public static IEnumerable<TSource> TakeLast<TSource>(this List<TSource> source, int count)
        {
            var sourceEnumerator = source.GetEnumerator();
            var buffer = new TSource[count];
            var numOfItems = 0;
            int idx;

            for (idx = 0; idx < count && sourceEnumerator.MoveNext(); idx++, numOfItems++)
            {
                buffer[idx] = sourceEnumerator.Current;
            }

            if (numOfItems < count)
            {
                for (idx = 0; idx < numOfItems; idx++)
                {
                    yield return buffer[idx];
                }

                yield break;
            }

            for (idx = 0; sourceEnumerator.MoveNext(); idx = (idx + 1) % count)
            {
                buffer[idx] = sourceEnumerator.Current;
            }

            for (; numOfItems > 0; idx = (idx + 1) % count, numOfItems--)
            {
                yield return buffer[idx];
            }
        }
    }
}