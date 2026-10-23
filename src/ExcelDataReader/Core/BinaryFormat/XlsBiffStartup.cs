using ExcelDataReader.Core.OfficeCrypto;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.BinaryFormat;

internal static class XlsBiffStartup
{
    public static XlsBiffStream Open(
        Stream source, out int version, out BIFFTYPE type, int offset = 0, int explicitVersion = 0, BIFFTYPE? defaultType = null, string? password = null, byte[]? secretKey = null, EncryptionInfo? encryption = null)
    {
        var stream = new XlsBiffStream(source, offset);
        try
        {
            var startup = Read(stream, explicitVersion, defaultType, password, secretKey, encryption);
            version = startup.Version;
            type = startup.Type;
            stream.InitializeEncryption(startup.SecretKey, startup.Encryption);
            stream.Position = offset;
            return stream;
        }
        catch
        {
            stream.Dispose();
            throw;
        }
    }

    private static (int Version, BIFFTYPE Type, byte[]? SecretKey, EncryptionInfo? Encryption) Read(
        XlsBiffStream stream, int explicitVersion, BIFFTYPE? defaultType, string? password, byte[]? secretKey, EncryptionInfo? encryption)
    {
        int version = 0;
        BIFFTYPE type = default;
        stream.ReadFrame(out var frame);
        if (frame.IsValid && frame.Id is BIFFRECORDTYPE.BOF or BIFFRECORDTYPE.BOF_V2 or BIFFRECORDTYPE.BOF_V3 or BIFFRECORDTYPE.BOF_V4)
        {
            version = explicitVersion == 0 ? GetVersion(frame) : explicitVersion;
            type = (BIFFTYPE)frame.ReadUInt16(2);
            if (secretKey == null)
                stream.ReadFrame(out frame);
        }
        else if (explicitVersion > 0 && defaultType != null)
        {
            version = explicitVersion;
            type = defaultType.Value;
        }

        if (secretKey != null)
            return (version, type, secretKey, encryption ?? throw new ArgumentNullException(nameof(encryption)));

        if (!frame.IsValid || frame.Id != BIFFRECORDTYPE.FILEPASS)
            stream.ReadFrame(out frame);

        if (frame.IsValid && frame.Id == BIFFRECORDTYPE.FILEPASS)
        {
            encryption = ReadEncryption(frame, version);
            if (encryption.VerifyPassword("VelvetSweatshop"))
                password = "VelvetSweatshop";
            else if (password == null || !encryption.VerifyPassword(password))
                throw new InvalidPasswordException(Errors.ErrorInvalidPassword);
            secretKey = encryption.GenerateSecretKey(password);
        }
        else
        {
            encryption = null;
        }

        return (version, type, secretKey, encryption);
    }

    private static int GetVersion(XlsBiffFrame frame) => frame.Id switch
    {
        BIFFRECORDTYPE.BOF_V2 => 2,
        BIFFRECORDTYPE.BOF_V3 => 3,
        BIFFRECORDTYPE.BOF_V4 => 4,
        _ => frame.ReadUInt16(0) switch
        {
            0x200 => 2,
            0x300 => 3,
            0x400 => 4,
            0x500 or 0 => 5,
            0x600 => 8,
            _ => 0,
        },
    };

    private static EncryptionInfo ReadEncryption(XlsBiffFrame frame, int version)
    {
        if (version is >= 2 and <= 5)
            return EncryptionInfo.Create(frame.ReadUInt16(0), frame.ReadUInt16(2));

        ushort type = frame.ReadUInt16(0);
        if (type == 0)
            return EncryptionInfo.Create(frame.ReadUInt16(2), frame.ReadUInt16(4));
        if (type != 1)
            throw new NotSupportedException("Unknown encryption type: " + type);

        var bytes = new byte[frame.RecordSize - 2];
        Buffer.BlockCopy(frame.Bytes, 6, bytes, 0, bytes.Length);
        return EncryptionInfo.Create(bytes);
    }
}