namespace ExcelDataReader.Core.OfficeCrypto
{
    internal enum CipherIdentifier
    {
        None,
        RC2,
        DES,
        DES3,
        AES,
        RC4
    }

    internal enum HashIdentifier
    {
        None,
        MD5,
        SHA1,
        SHA256,
        SHA384,
        SHA512,
    }
}
