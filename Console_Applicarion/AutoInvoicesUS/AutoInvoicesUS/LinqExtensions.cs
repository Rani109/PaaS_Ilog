using System;
using System.Collections.Generic;

namespace System.Linq
{
    public static partial class LinqExtensions
    {
        /// <summary>
        /// ForEach
        /// </summary>
        public static IEnumerable<TSource> ForEach<TSource>(this IEnumerable<TSource> source, Action<TSource> action)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            if (action == null)
                throw new ArgumentNullException("action");

            foreach (TSource item in source)
                action(item);

            return source;
        }

        /// <summary>
        /// Breakable ForEach
        /// </summary>
        public static IEnumerable<TSource> ForEach<TSource>(this IEnumerable<TSource> source, Func<TSource, bool> action)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            if (action == null)
                throw new ArgumentNullException("action");

            foreach (TSource item in source)
                if (!action(item))
                    break;

            return source;
        }

        /// <summary>
        /// ForEach with index
        /// </summary>
        public static IEnumerable<TSource> ForEach<TSource>(this IEnumerable<TSource> source, Action<TSource, int> action)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            if (action == null)
                throw new ArgumentNullException("action");

            int index = 0;
            foreach (TSource item in source)
                action(item, index++);

            return source;
        }

        /// <summary>
        /// Breakable ForEach with index
        /// </summary>
        public static IEnumerable<TSource> ForEach<TSource>(this IEnumerable<TSource> source, Func<TSource, int, bool> action)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            if (action == null)
                throw new ArgumentNullException("action");

            int index = 0;
            foreach (TSource item in source)
                if (!action(item, index++))
                    break;

            return source;
        }

        /// <summary>
        /// ForEach with index and previous item
        /// </summary>
        public static IEnumerable<TSource> ForEach<TSource>(this IEnumerable<TSource> source, Action<TSource, int, TSource> action)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            if (action == null)
                throw new ArgumentNullException("action");

            int index = 0;
            TSource previousItem = default(TSource);
            foreach (TSource item in source)
            {
                action(item, index++, previousItem);
                previousItem = item;
            }

            return source;
        }

        /// <summary>
        /// Breakable ForEach with index and previous item
        /// </summary>
        public static IEnumerable<TSource> ForEach<TSource>(this IEnumerable<TSource> source, Func<TSource, int, TSource, bool> action)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            if (action == null)
                throw new ArgumentNullException("action");

            int index = 0;
            TSource previousItem = default(TSource);
            foreach (TSource item in source)
            {
                if (!action(item, index++, previousItem))
                    break;
                previousItem = item;
            }

            return source;
        }

        /// <summary>
        /// ForEach with index and 2 previous items
        /// </summary>
        public static IEnumerable<TSource> ForEach<TSource>(this IEnumerable<TSource> source, Action<TSource, int, TSource, TSource> action)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            if (action == null)
                throw new ArgumentNullException("action");

            int index = 0;
            TSource previousItem1 = default(TSource);
            TSource previousItem2 = default(TSource);
            foreach (TSource item in source)
            {
                action(item, index++, previousItem1, previousItem2);
                previousItem2 = previousItem1;
                previousItem1 = item;
            }

            return source;
        }

        /// <summary>
        /// Breakable ForEach with index and 2 previous items
        /// </summary>
        public static IEnumerable<TSource> ForEach<TSource>(this IEnumerable<TSource> source, Func<TSource, int, TSource, TSource, bool> action)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            if (action == null)
                throw new ArgumentNullException("action");

            int index = 0;
            TSource previousItem1 = default(TSource);
            TSource previousItem2 = default(TSource);
            foreach (TSource item in source)
            {
                if (!action(item, index++, previousItem1, previousItem2))
                    break;
                previousItem2 = previousItem1;
                previousItem1 = item;
            }

            return source;
        }

        /// <summary>
        /// ForEach with index and previous result
        /// </summary>
        public static IEnumerable<TSource> ForEach<TSource, TResult>(this IEnumerable<TSource> source, Func<TSource, int, TResult, TResult> action)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            if (action == null)
                throw new ArgumentNullException("action");

            int index = 0;
            TResult previousResult = default(TResult);
            foreach (TSource item in source)
                previousResult = action(item, index++, previousResult);

            return source;
        }

        /// <summary>
        /// ForEach with index and 2 previous results
        /// </summary>
        public static IEnumerable<TSource> ForEach<TSource, TResult>(this IEnumerable<TSource> source, Func<TSource, int, TResult, TResult, TResult> action)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            if (action == null)
                throw new ArgumentNullException("action");

            int index = 0;
            TResult previousResult1 = default(TResult);
            TResult previousResult2 = default(TResult);
            foreach (TSource item in source)
            {
                TResult result = action(item, index++, previousResult1, previousResult2);
                previousResult2 = previousResult1;
                previousResult1 = result;
            }

            return source;
        }
    }
}