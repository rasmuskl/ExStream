using System;
using System.Collections.Generic;
using ExStream.CellWriters;

namespace ExStream
{
    public class ExStreamWriterConfig
    {
        public static readonly ExStreamWriterConfig Current = new ExStreamWriterConfig();

        readonly Dictionary<Type, ICellWriter> _cellWriters = new Dictionary<Type, ICellWriter>();

        public ExStreamWriterConfig()
        {
            AddCellWriter(typeof(string), new StringCellWriter());
            AddCellWriter(typeof(Guid), new StringCellWriter());
            AddCellWriter(typeof(DateTime), new DateTimeCellWriter());
            AddCellWriter(typeof(bool), new BooleanCellWriter());

            var numberTypes = new[]
            {
                typeof (short),
                typeof (int),
                typeof (long),
                typeof (byte),
                typeof (ushort),
                typeof (uint),
                typeof (ulong),
                typeof (double),
                typeof (float),
                typeof (decimal),
            };

            foreach (var numberType in numberTypes)
            {
                AddCellWriter(numberType, new NumberCellWriter());
            }
        }

        public void ClearCellWriters()
        {
            _cellWriters.Clear();
        }

        public void AddCellWriter(Type type, ICellWriter cellWriter)
        {
            _cellWriters[type] = cellWriter;
        }

        public ICellWriter GetCellWriter(Type type)
        {
            ICellWriter writer;

            if (_cellWriters.TryGetValue(type, out writer))
            {
                return writer;
            }

            return null;
        }
    }
}