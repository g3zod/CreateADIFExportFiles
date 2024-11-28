using System;
#if INCLUDE_TEST_CODE
using System.Diagnostics;
#endif

namespace AdifReleaseLib
{
    /**
     * <summary>
     *   Contains extension methods.
     * </summary>
     */
    public static class Extensions
    {
        /**
         * <summary>
         *   Compares two strings that contain lines separated by "\r\n".
         * </summary>
         * 
         * <param name="source">The instance <see cref="string"/>.</param>
         * <param name="comparison">The <see cref="string"/> to be compared with the instance <see cref="string"/>.</param>
         * 
         * <returns>A <see cref="string"/> that is either zero-length or contains details of the first discrepancy between the strings.</returns>
         */
        public static string CompareLines(this string source, string comparison)
        {
            // Because CompareLines is called as an instance method, I can't see how sourceStr can be null.
            // However, err on the side of caution and check anyway.

            ArgumentNullException.ThrowIfNull(source);
            ArgumentNullException.ThrowIfNull(comparison);

            string result = string.Empty;

            if (source != comparison)
            {
                string[]
                    sourceLines = source.Split("\r\n"),
                    comparisonLines = comparison.Split("\r\n");

                for (int i = 0; i < int.Max(sourceLines.Length, comparisonLines.Length); i++)
                {
                    if (i == sourceLines.Length)
                    {
                        result = $"Line {i + 1} does not exist in source string: \"{comparisonLines[i]}\"";
                        break;
                    }
                    else if (i == comparisonLines.Length)
                    {
                        result = $"Line {i + 1} does not exist in comparison string: \"{sourceLines[i]}\"";
                        break;
                    }
                    else
                    {
                        string
                            sourceLine = sourceLines[i],
                            comparisonLine = comparisonLines[i];

                        if (comparisonLine != sourceLine)
                        {
                            result = $"Line {i + 1} is different: source=\"{sourceLine}\", comparison=\"{comparisonLine}\"";
                            break;
                        }
                    }
                }
            }
            return result;
        }

        /**
         * <summary>
         *   Truncates a <see cref="DateTime"/> to a specified resolution.
         *
         * <para>
         *   <example>
         *     For example, to truncate to a whole number of seconds:
         *     <code>DateTime dt = DateTime.Now.Truncate(TimeSpan.TicksPerSecond);</code>
         *   </example>
         * </para>
         * </summary>
         * 
         * <param name="dateTime">The <see cref="DateTime"/> object to truncate</param>
         * <param name="resolution">
         *   The number of Ticks to truncate to, e.g. for days, <see cref="TimeSpan.TicksPerDay"/>.
         *   The default value is <see cref="TimeSpan.TicksPerSecond"/>.
         * </param>
         * 
         * <returns>Truncated DateTime.</returns>
         *
         * <remarks>Based on <see cref="DateTime"/><see href="https://stackoverflow.com/a/2681758"/></remarks>
         */
        public static DateTime Truncate(this DateTime dateTime, long resolution = TimeSpan.TicksPerSecond)
        {
            return new DateTime(dateTime.Ticks - (dateTime.Ticks % resolution), dateTime.Kind);
        }

#if INCLUDE_TEST_CODE
        /**
         * <summary>
         *   A method that can be called to test the <see cref="CompareLines(string, string)"/> extension method.
         * </summary>
         */
        public static void TestExtensionMethods()
        {
            {
                string test;

                test = "abc\r\ndef".CompareLines("abc\r\ndef\r\nghi");
                Trace.Assert(test == "Line 3 does not exist in source string: \"ghi\"");

                test = "abc\r\ndef\r\nghi".CompareLines("abc\r\ndef");
                Trace.Assert(test == "Line 3 does not exist in comparison string: \"ghi\"");

                test = "abc\r\ndef\r\nghi".CompareLines("123\r\ndef\r\nghi");
                Trace.Assert(test == "Line 1 is different: source=\"abc\", comparison=\"123\"");

                test = "abc\r\ndef\r\nghi".CompareLines("abc\r\n456\r\nghi");
                Trace.Assert(test == "Line 2 is different: source=\"def\", comparison=\"456\"");

                test = "abc\r\ndef\r\nghi".CompareLines("abc\r\ndef\r\n789");
                Trace.Assert(test == "Line 3 is different: source=\"ghi\", comparison=\"789\"");
            }

            {
                DateTime
                    source = DateTime.UtcNow,
                    expectedSeconds = new(
                        source.Year,
                        source.Month,
                        source.Day,
                        source.Hour,
                        source.Minute,
                        source.Second,
                        DateTimeKind.Utc),
                    expectedDate = new(
                        source.Year,
                        source.Month,
                        source.Day,
                        00,
                        00,
                        00,
                        DateTimeKind.Utc);

                Trace.Assert(source.Truncate() == expectedSeconds);
                Trace.Assert(source.Truncate(TimeSpan.TicksPerSecond) == expectedSeconds);
                Trace.Assert(source.Truncate(TimeSpan.TicksPerDay) == expectedDate);
                Trace.Assert(source.Truncate().Kind == DateTimeKind.Utc);
            }
        }
#endif
    }
}
