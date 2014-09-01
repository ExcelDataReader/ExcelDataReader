using System.Threading.Tasks;
using PCLStorage;

namespace ExcelDataReader.Portable.IO.PCLStorage
{
    public static class FolderExtensions
    {
        /// <summary>
        /// Deletes the folder and all its contents
        /// </summary>
        /// <param name="folder"></param>
        /// <returns></returns>
        public static async Task DeleteFolderAndContentsAsync(this IFolder folder)
        {
            var subFolders = await folder.GetFoldersAsync();

            //recursively delete all subfolders
            foreach (var subFolder in subFolders)
            {
                await subFolder.DeleteFolderAndContentsAsync();
            }
            //delete all files
            var files = await folder.GetFilesAsync();
            foreach (var file in files)
            {
                await file.DeleteAsync();
            }
            //delete the folder

            await folder.DeleteAsync();
        }
    }
}
