using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.CompilerServices;
using System.Xml.Linq;

namespace FakeExcelSerializer;

public class ArrayPoolBufferWriter : IBufferWriter<byte>, IDisposable
{
    byte[]? _rentedBuffer;
    int _written;
    long _committed;

    const int MinimumBufferSize = 256;

    public ArrayPoolBufferWriter(int initialCapacity = MinimumBufferSize)
    {
        if (initialCapacity <= 0)
            throw new ArgumentException(null, nameof(initialCapacity));

        _rentedBuffer = ArrayPool<byte>.Shared.Rent(initialCapacity);
        _written = 0;
        _committed = 0;
    }

    public Memory<byte> OutputAsMemory
    {
        get
        {
            CheckIfDisposed();
            return _rentedBuffer.AsMemory(0, _written);
        }
    }

    public Span<byte> OutputAsSpan
    {
        get
        {
            CheckIfDisposed();
            return _rentedBuffer.AsSpan(0, _written);
        }
    }

    public int BytesWritten
    {
        get
        {
            CheckIfDisposed();
            return _written;
        }
    }

    public long BytesCommitted
    {
        get
        {
            CheckIfDisposed();
            return _committed;
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Clear()
    {
        CheckIfDisposed();
        ClearHelper();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void ClearHelper()
    {
        _rentedBuffer.AsSpan(0, _written).Clear();
        _written = 0;
    }

    public async Task CopyToAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        CheckIfDisposed();

        if (stream == null)
            ThrowArgumentNullException(nameof(stream));
        if (_rentedBuffer == null)
            ThrowArgumentNullException(nameof(_rentedBuffer));

        await stream.WriteAsync(_rentedBuffer, 0, _written, cancellationToken).ConfigureAwait(false);
        _committed += _written;

        ClearHelper();
    }

    public void CopyTo(Stream stream)
    {
        CheckIfDisposed();

        if (stream == null)
            ThrowArgumentNullException(nameof(stream));

        if (_rentedBuffer == null)
            ThrowArgumentNullException(nameof(_rentedBuffer));

        stream.Write(_rentedBuffer, 0, _written);
        _committed += _written;

        ClearHelper();
    }

    public void Advance(int count)
    {
        CheckIfDisposed();

        if (count < 0)
            ThrowArgumentNullException(nameof(count));

        if (_written > (_rentedBuffer?.Length ?? 0) - count)
            ThrowInvalidOperationException();

        _written += count;
    }

    public void Dispose()
    {
        if (_rentedBuffer == null)
            return;

        ArrayPool<byte>.Shared.Return(_rentedBuffer, clearArray: true);
        _rentedBuffer = null;
        _written = 0;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void CheckIfDisposed()
    {
        if (_rentedBuffer == null)
            ThrowObjectDisposedException();
    }

    public Memory<byte> GetMemory(int sizeHint = 0)
    {
        CheckIfDisposed();

        if (sizeHint < 0)
            ThrowArgumentException(nameof(sizeHint));

        CheckAndResizeBuffer(sizeHint);
        return _rentedBuffer.AsMemory(_written);
    }

    public Span<byte> GetSpan(int sizeHint = 0)
    {
        CheckIfDisposed();

        if (sizeHint < 0)
            ThrowArgumentException(nameof(sizeHint));

        CheckAndResizeBuffer(sizeHint);
        return _rentedBuffer.AsSpan(_written);
    }

    void CheckAndResizeBuffer(int sizeHint)
    {
        if (_rentedBuffer == null)
        {
            _rentedBuffer = ArrayPool<byte>.Shared.Rent(Math.Max(MinimumBufferSize, sizeHint));
            return;
        }

        if (sizeHint == 0)
            sizeHint = MinimumBufferSize;

        int availableSpace = _rentedBuffer.Length - _written;

        if (sizeHint > availableSpace)
        {
            int growBy = sizeHint > _rentedBuffer.Length ? sizeHint : _rentedBuffer.Length;

            int newSize = checked(_rentedBuffer.Length + growBy);

            byte[] oldBuffer = _rentedBuffer;

            _rentedBuffer = ArrayPool<byte>.Shared.Rent(newSize);

            oldBuffer.AsSpan(0, _written).CopyTo(_rentedBuffer);
            ArrayPool<byte>.Shared.Return(oldBuffer, clearArray: true);
        }
    }

    [DoesNotReturn]
    static void ThrowArgumentNullException(string name)
        => throw new ArgumentNullException(name);
    [DoesNotReturn]
    static void ThrowObjectDisposedException()
        => throw new ObjectDisposedException(nameof(ArrayPoolBufferWriter));
    [DoesNotReturn]
    static void ThrowInvalidOperationException()
        => throw new InvalidOperationException("Cannot advance past the end of the buffer.");
    [DoesNotReturn]
    static void ThrowArgumentException(string name)
        => throw new ArgumentException(null, name);
}

