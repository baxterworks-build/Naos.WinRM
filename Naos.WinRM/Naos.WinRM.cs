﻿namespace Naos.WinRM.Core
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Globalization;
    using System.Linq;
    using System.Management.Automation;
    using System.Management.Automation.Runspaces;
    using System.Runtime.Serialization;
    using System.Security;
    using System.Security.Cryptography;

    /// <summary>
    /// Custom base exception to allow global catching of internally generated errors.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Rm", Justification = "Name/spelling is correct.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Rm", Justification = "Name/spelling is correct.")]
    [Serializable]
    public abstract class NaosWinRmBaseException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="NaosWinRmBaseException"/> class.
        /// </summary>
        protected NaosWinRmBaseException()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="NaosWinRmBaseException"/> class.
        /// </summary>
        /// <param name="message">Exception message.</param>
        protected NaosWinRmBaseException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="NaosWinRmBaseException"/> class.
        /// </summary>
        /// <param name="message">Exception message.</param>
        /// <param name="innerException">Inner exception.</param>
        protected NaosWinRmBaseException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="NaosWinRmBaseException"/> class.
        /// </summary>
        /// <param name="info">Serialization info.</param>
        /// <param name="context">Reading context.</param>
        protected NaosWinRmBaseException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }

    /// <summary>
    /// Custom exception for when trying to execute 
    /// </summary>
    [Serializable]
    public class TrustedHostMissingException : NaosWinRmBaseException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TrustedHostMissingException"/> class.
        /// </summary>
        public TrustedHostMissingException()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TrustedHostMissingException"/> class.
        /// </summary>
        /// <param name="message">Exception message.</param>
        public TrustedHostMissingException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TrustedHostMissingException"/> class.
        /// </summary>
        /// <param name="message">Exception message.</param>
        /// <param name="innerException">Inner exception.</param>
        public TrustedHostMissingException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TrustedHostMissingException"/> class.
        /// </summary>
        /// <param name="info">Serialization info.</param>
        /// <param name="context">Reading context.</param>
        protected TrustedHostMissingException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }

    /// <summary>
    /// Custom exception for when things go wrong running remote commands.
    /// </summary>
    [Serializable]
    public class RemoteExecutionException : NaosWinRmBaseException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RemoteExecutionException"/> class.
        /// </summary>
        public RemoteExecutionException()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RemoteExecutionException"/> class.
        /// </summary>
        /// <param name="message">Exception message.</param>
        public RemoteExecutionException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RemoteExecutionException"/> class.
        /// </summary>
        /// <param name="message">Exception message.</param>
        /// <param name="innerException">Inner exception.</param>
        public RemoteExecutionException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RemoteExecutionException"/> class.
        /// </summary>
        /// <param name="info">Serialization info.</param>
        /// <param name="context">Reading context.</param>
        protected RemoteExecutionException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }

    /// <summary>
    /// Manages various remote tasks on a machine using the WinRM protocol.
    /// </summary>
    public interface IManageMachines
    {
        /// <summary>
        /// Gets the IP address of the machine being managed.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Ip", Justification = "Name/spelling is correct.")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Ip", Justification = "Name/spelling is correct.")]
        string IpAddress { get; }

        /// <summary>
        /// Executes a user initiated reboot.
        /// </summary>
        /// <param name="force">Can override default behavior of a forceful reboot (kick users off).</param>
        void Reboot(bool force = true);

        /// <summary>
        /// Sends a file to the remote machine at the provided file path on that target computer.
        /// </summary>
        /// <param name="filePathOnTargetMachine">File path to write the contents to on the remote machine.</param>
        /// <param name="fileContents">Payload to write to the file.</param>
        /// <param name="appended">Optionally writes the bytes in appended mode or not (default is NOT).</param>
        /// <param name="overwrite">Optionally will overwrite a file that is already there [can NOT be used with 'appended'] (default is NOT).</param>
        void SendFile(string filePathOnTargetMachine, byte[] fileContents, bool appended = false, bool overwrite = false);

        /// <summary>
        /// Retrieves a file from the remote machines and returns a checksum verified byte array.
        /// </summary>
        /// <param name="filePathOnTargetMachine">File path to fetch the contents of on the remote machine.</param>
        /// <returns>Bytes of the specified files (throws if missing).</returns>
        byte[] RetrieveFile(string filePathOnTargetMachine);

        /// <summary>
        /// Runs an arbitrary command using "CMD.exe /c".
        /// </summary>
        /// <param name="command">Command to run in "CMD.exe".</param>
        /// <param name="commandParameters">Parameters to be passed to the command.</param>
        /// <returns>Console output of the command.</returns>
        string RunCmd(string command, ICollection<string> commandParameters = null);

        /// <summary>
        /// Runs an arbitrary command using "CMD.exe /c" on localhost instead of the provided remote computer..
        /// </summary>
        /// <param name="command">Command to run in "CMD.exe".</param>
        /// <param name="commandParameters">Parameters to be passed to the command.</param>
        /// <returns>Console output of the command.</returns>
        string RunCmdOnLocalhost(string command, ICollection<string> commandParameters = null);

        /// <summary>
        /// Runs an arbitrary script block on localhost instead of the provided remote computer.
        /// </summary>
        /// <param name="scriptBlock">Script block.</param>
        /// <param name="scriptBlockParameters">Parameters to be passed to the script block.</param>
        /// <returns>Collection of objects that were the output from the script block.</returns>
        ICollection<dynamic> RunScriptOnLocalhost(string scriptBlock, ICollection<object> scriptBlockParameters = null);

        /// <summary>
        /// Runs an arbitrary script block.
        /// </summary>
        /// <param name="scriptBlock">Script block.</param>
        /// <param name="scriptBlockParameters">Parameters to be passed to the script block.</param>
        /// <returns>Collection of objects that were the output from the script block.</returns>
        ICollection<dynamic> RunScript(string scriptBlock, ICollection<object> scriptBlockParameters = null);
    }

    /// <inheritdoc />
    public class MachineManager : IManageMachines
    {
        private readonly long fileChunkSizeThresholdByteCount;

        private readonly long fileChunkSizePerSend;

        private readonly long fileChunkSizePerRetrieve;

        private readonly string userName;

        private readonly SecureString password;

        private readonly bool autoManageTrustedHosts;

        private static readonly object SyncTrustedHosts = new object();

        /// <summary>
        /// Initializes a new instance of the <see cref="MachineManager"/> class.
        /// </summary>
        /// <param name="ipAddress">IP address of machine to interact with.</param>
        /// <param name="userName">Username to use to connect.</param>
        /// <param name="password">Password to use to connect.</param>
        /// <param name="autoManageTrustedHosts">Optionally specify whether to update the TrustedHost list prior to execution or assume it's handled elsewhere (default is FALSE).</param>
        /// <param name="fileChunkSizeThresholdByteCount">Optionally specify file size that will trigger chunking the file rather than sending as one file (150000 is default).</param>
        /// <param name="fileChunkSizePerSend">Optionally specify size of each chunk that is sent when a file is being chunked for send.</param>
        /// <param name="fileChunkSizePerRetrieve">Optionally specify size of each chunk that is received when a file is being chunked for fetch.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "ip", Justification = "Name/spelling is correct.")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "byte", Justification = "Name/spelling is correct.")]
        public MachineManager(
            string ipAddress,
            string userName,
            SecureString password,
            bool autoManageTrustedHosts = false,
            long fileChunkSizeThresholdByteCount = 150000,
            long fileChunkSizePerSend = 100000,
            long fileChunkSizePerRetrieve = 100000)
        {
            this.IpAddress = ipAddress;
            this.userName = userName;
            this.password = password;
            this.autoManageTrustedHosts = autoManageTrustedHosts;
            this.fileChunkSizeThresholdByteCount = fileChunkSizeThresholdByteCount;
            this.fileChunkSizePerSend = fileChunkSizePerSend;
            this.fileChunkSizePerRetrieve = fileChunkSizePerRetrieve;
        }

        /// <summary>
        /// Locally updates the trusted hosts to have the ipAddress provided.
        /// </summary>
        /// <param name="ipAddress">IP Address to add to local trusted hosts.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "ip", Justification = "Name/spelling is correct.")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Ip", Justification = "Name/spelling is correct.")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Ip", Justification = "Name/spelling is correct.")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId = "notUsedOutput", Justification = "Prefer to see that output is generated and not used...")]
        public static void AddIpAddressToLocalTrustedHosts(string ipAddress)
        {
            lock (SyncTrustedHosts)
            {
                var currentTrustedHosts = GetListOfIpAddressesFromLocalTrustedHosts().ToList();

                if (!currentTrustedHosts.Contains(ipAddress) && !TrustedHostListIsWildCard(currentTrustedHosts))
                {
                    currentTrustedHosts.Add(ipAddress);
                    var newValue = currentTrustedHosts.Any() ? string.Join(",", currentTrustedHosts) : ipAddress;
                    using (var runspace = RunspaceFactory.CreateRunspace())
                    {
                        runspace.Open();

                        var command = new Command("Set-Item");
                        command.Parameters.Add("Path", @"WSMan:\localhost\Client\TrustedHosts");
                        command.Parameters.Add("Value", newValue);
                        command.Parameters.Add("Force", true);

                        var notUsedOutput = RunLocalCommand(runspace, command);
                    }
                }
            }
        }

        /// <summary>
        /// Locally updates the trusted hosts to remove the ipAddress provided (if applicable).
        /// </summary>
        /// <param name="ipAddress">IP Address to remove from local trusted hosts.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "ip", Justification = "Name/spelling is correct.")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Ip", Justification = "Name/spelling is correct.")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Ip", Justification = "Name/spelling is correct.")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId = "notUsedOutput", Justification = "Prefer to see that output is generated and not used...")]
        public static void RemoveIpAddressFromLocalTrustedHosts(string ipAddress)
        {
            lock (SyncTrustedHosts)
            {
                var currentTrustedHosts = GetListOfIpAddressesFromLocalTrustedHosts().ToList();

                if (currentTrustedHosts.Contains(ipAddress))
                {
                    currentTrustedHosts.Remove(ipAddress);

                    // can't pass null must be an empty string...
                    var newValue = currentTrustedHosts.Any() ? string.Join(",", currentTrustedHosts) : string.Empty;

                    using (var runspace = RunspaceFactory.CreateRunspace())
                    {
                        runspace.Open();

                        var command = new Command("Set-Item");
                        command.Parameters.Add("Path", @"WSMan:\localhost\Client\TrustedHosts");
                        command.Parameters.Add("Value", newValue);
                        command.Parameters.Add("Force", true);

                        var notUsedOutput = RunLocalCommand(runspace, command);
                    }
                }
            }
        }

        /// <summary>
        /// Locally updates the trusted hosts to have the ipAddress provided.
        /// </summary>
        /// <returns>List of the trusted hosts.</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Justification = "Want a method due to amount of logic.")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Ip", Justification = "Name/spelling is correct.")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Ip", Justification = "Name/spelling is correct.")]
        public static IReadOnlyCollection<string> GetListOfIpAddressesFromLocalTrustedHosts()
        {
            lock (SyncTrustedHosts)
            {
                try
                {
                    using (var runspace = RunspaceFactory.CreateRunspace())
                    {
                        runspace.Open();

                        var command = new Command("Get-Item");
                        command.Parameters.Add("Path", @"WSMan:\localhost\Client\TrustedHosts");

                        var response = RunLocalCommand(runspace, command);

                        var valueProperty = response.Single().Properties.Single(_ => _.Name == "Value");

                        var value = valueProperty.Value.ToString();

                        var ret = string.IsNullOrEmpty(value) ? new string[0] : value.Split(',');

                        return ret;
                    }
                }
                catch (RemoteExecutionException remoteException)
                {
                    // if we don't have any trusted hosts then just ignore...
                    if (
                        remoteException.Message.Contains(
                            "Cannot find path 'WSMan:\\localhost\\Client\\TrustedHosts' because it does not exist."))
                    {
                        return new List<string>();
                    }

                    throw;
                }
            }
        }

        /// <inheritdoc />
        public string IpAddress { get; private set; }

        /// <inheritdoc />
        public void Reboot(bool force = true)
        {
            var forceAddIn = force ? " -Force" : string.Empty;
            var restartScriptBlock = "{ Restart-Computer" + forceAddIn + " }";
            this.RunScript(restartScriptBlock);
        }

        /// <inheritdoc />
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times", Justification = "Disposal logic is correct.")]
        public void SendFile(string filePathOnTargetMachine, byte[] fileContents, bool appended = false, bool overwrite = false)
        {
            if (fileContents == null)
            {
                throw new ArgumentNullException(nameof(fileContents));
            }

            if (appended && overwrite)
            {
                throw new ArgumentException("Cannot run with overwrite AND appended.");
            }

            using (var runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();

                var sessionObject = this.BeginSession(runspace);

                var verifyFileDoesntExistScriptBlock = @"
	                { 
		                param($filePath)

			            if (Test-Path $filePath)
			            {
				            throw ""File already exists at: $filePath""
			            }
	                }";

                if (!appended && !overwrite)
                {
                    this.RunScriptUsingSession(
                        verifyFileDoesntExistScriptBlock,
                        new[] { filePathOnTargetMachine },
                        runspace,
                        sessionObject);
                }

                var firstSendUsingSession = true;
                if (fileContents.Length <= this.fileChunkSizeThresholdByteCount)
                {
                    this.SendFileUsingSession(filePathOnTargetMachine, fileContents, appended, overwrite, runspace, sessionObject);
                }
                else
                {
                    // deconstruct and send pieces as appended...
                    var nibble = new List<byte>();
                    foreach (byte currentByte in fileContents)
                    {
                        if (nibble.Count < this.fileChunkSizePerSend)
                        {
                            nibble.Add(currentByte);
                        }
                        else
                        {
                            nibble.Add(currentByte);
                            this.SendFileUsingSession(filePathOnTargetMachine, nibble.ToArray(), !firstSendUsingSession, overwrite && firstSendUsingSession, runspace, sessionObject);
                            firstSendUsingSession = false;
                            nibble.Clear();
                        }
                    }

                    // flush the "buffer"...
                    if (nibble.Any())
                    {
                        this.SendFileUsingSession(filePathOnTargetMachine, nibble.ToArray(), true, false, runspace, sessionObject);
                        nibble.Clear();
                    }
                }

                var expectedChecksum = ComputeSha256Hash(fileContents);
                var verifyChecksumScriptBlock = @"
	                { 
		                param($filePath, $expectedChecksum)

		                $fileToCheckFileInfo = New-Object System.IO.FileInfo($filePath)
		                if (-not $fileToCheckFileInfo.Exists)
		                {
			                # If the file can't be found, try looking for it in the current directory.
			                $fileToCheckFileInfo = New-Object System.IO.FileInfo($filePath)
			                if (-not $fileToCheckFileInfo.Exists)
			                {
				                throw ""Can't find the file specified to calculate a checksum on: $filePath""
			                }
		                }

		                $fileToCheckFileStream = $fileToCheckFileInfo.OpenRead()
                        $provider = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider
                        $hashBytes = $provider.ComputeHash($fileToCheckFileStream)
		                $fileToCheckFileStream.Close()
		                $fileToCheckFileStream.Dispose()
		
		                $base64 = [System.Convert]::ToBase64String($hashBytes)
		
		                $calculatedChecksum = [System.String]::Empty
		                foreach ($byte in $hashBytes)
		                {
			                $calculatedChecksum = $calculatedChecksum + $byte.ToString(""X2"")
		                }

		                if($calculatedChecksum -ne $expectedChecksum)
		                {
			                Write-Error ""Checksums don't match on File: $filePath - Expected: $expectedChecksum - Actual: $calculatedChecksum""
		                }
	                }";

                this.RunScriptUsingSession(
                    verifyChecksumScriptBlock,
                    new[] { filePathOnTargetMachine, expectedChecksum },
                    runspace,
                    sessionObject);

                this.EndSession(sessionObject, runspace);

                runspace.Close();
            }
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId = "notUsedResults", Justification = "Prefer to see that output is generated and not used...")]
        private void SendFileUsingSession(
            string filePathOnTargetMachine,
            byte[] fileContents,
            bool appended,
            bool overwrite,
            Runspace runspace,
            object sessionObject)
        {
            if (appended && overwrite)
            {
                throw new ArgumentException("Cannot run with overwrite AND appended.");
            }

            var commandName = appended ? "Add-Content" : "Set-Content";
            var forceAddIn = overwrite ? " -Force" : string.Empty;
            var sendFileScriptBlock = @"
	                { 
		                param($filePath, $fileContents)

		                $parentDir = Split-Path $filePath
		                if (-not (Test-Path $parentDir))
		                {
			                md $parentDir | Out-Null
		                }

		                " + commandName + @" -Path $filePath -Encoding Byte -Value $fileContents" + forceAddIn + @"
	                }";

            var arguments = new object[] { filePathOnTargetMachine, fileContents };

            var notUsedResults = this.RunScriptUsingSession(sendFileScriptBlock, arguments, runspace, sessionObject);
        }

        /// <inheritdoc />
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times", Justification = "Disposal logic is correct.")]
        public byte[] RetrieveFile(string filePathOnTargetMachine)
        {
            using (var runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();

                var sessionObject = this.BeginSession(runspace);

                var verifyFileExistsScriptBlock = @"
	                { 
		                param($filePath)

			            if (-not (Test-Path $filePath))
			            {
				            throw ""File doesn't exist at: $filePath""
			            }

                        $file = ls $filePath
                        Write-Output $file.Length
	                }";

                var fileSizeRaw = this.RunScriptUsingSession(
                    verifyFileExistsScriptBlock,
                    new[] { filePathOnTargetMachine },
                    runspace,
                    sessionObject);

                var fileSize = (long)long.Parse(fileSizeRaw.Single().ToString());

                var getChecksumScriptBlock = @"
	                { 
		                param($filePath)

		                $fileToCheckFileInfo = New-Object System.IO.FileInfo($filePath)
		                if (-not $fileToCheckFileInfo.Exists)
		                {
			                # If the file can't be found, try looking for it in the current directory.
			                $fileToCheckFileInfo = New-Object System.IO.FileInfo($filePath)
			                if (-not $fileToCheckFileInfo.Exists)
			                {
				                throw ""Can't find the file specified to calculate a checksum on: $filePath""
			                }
		                }

		                $fileToCheckFileStream = $fileToCheckFileInfo.OpenRead()
                        $provider = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider
                        $hashBytes = $provider.ComputeHash($fileToCheckFileStream)
		                $fileToCheckFileStream.Close()
		                $fileToCheckFileStream.Dispose()
		
		                $base64 = [System.Convert]::ToBase64String($hashBytes)
		
		                $calculatedChecksum = [System.String]::Empty
		                foreach ($byte in $hashBytes)
		                {
			                $calculatedChecksum = $calculatedChecksum + $byte.ToString(""X2"")
		                }
		                
                                # trimming off leading and trailing curly braces '{ }'
		                $trimmedChecksum = $calculatedChecksum.Substring(1, $calculatedChecksum.Length - 2)

                        Write-Output $trimmedChecksum
	                }";

                var remoteChecksumRaw = this.RunScriptUsingSession(
                    getChecksumScriptBlock,
                    new[] { filePathOnTargetMachine },
                    runspace,
                    sessionObject);

                var remoteChecksum = remoteChecksumRaw.Single();

                var bytes = new List<byte>();
                if (fileSize <= this.fileChunkSizeThresholdByteCount)
                {
                    var bytesRaw = this.RetrieveFileUsingSession(filePathOnTargetMachine, runspace, sessionObject);
                    bytes.AddRange(bytesRaw);
                }
                else
                {
                    // deconstruct and fetch pieces...
                    var lastNibblePoint = 0;
                    for (var nibblePoint = 0; nibblePoint < fileSize; nibblePoint++)
                    {
                        if ((nibblePoint - lastNibblePoint) >= this.fileChunkSizePerRetrieve)
                        {
                            var remainingBytes = fileSize - nibblePoint;
                            var nibbleSize = remainingBytes < this.fileChunkSizePerRetrieve
                                                 ? remainingBytes
                                                 : this.fileChunkSizePerRetrieve;

                            var nibble = this.RetrieveFileUsingSession(
                                filePathOnTargetMachine,
                                runspace,
                                sessionObject,
                                lastNibblePoint,
                                nibbleSize);
                            bytes.AddRange(nibble);
                            lastNibblePoint = nibblePoint;
                        }
                    }
                }

                var byteArray = bytes.ToArray();
                var actualChecksum = ComputeSha256Hash(byteArray);
                if (string.Equals(remoteChecksum.ToString(), actualChecksum.ToString(), StringComparison.CurrentCultureIgnoreCase))
                {
                    throw new RemoteExecutionException("Checksum didn't match after file was downloaded.");
                }

                this.EndSession(sessionObject, runspace);

                runspace.Close();

                return byteArray;
            }
        }

        private byte[] RetrieveFileUsingSession(string filePathOnTargetMachine, Runspace runspace, object sessionObject, long nibbleStart = 0, long nibbleSize = 0)
        {
            if (nibbleStart != 0 && nibbleSize == 0)
            {
                nibbleSize = this.fileChunkSizePerRetrieve;
            }

            var fetchFileScriptBlock = @"
	                { 
		                param($filePath, $nibbleStart, $nibbleSize)

		                if (-not (Test-Path $filePath))
		                {
			                throw ""Expected file to fetch missing at: $filePath""
		                }

                        $allBytes = [System.IO.File]::ReadAllBytes($filePath)
                        if (($nibbleStart -eq 0) -and ($nibbleSize -eq 0))
                        {
                            Write-Output $allBytes
                        }
                        else
                        {
                            $nibble = new-object byte[] $nibbleSize
                            [Array]::Copy($allBytes, $nibbleStart, $nibble, 0, $nibbleSize)
                            Write-Output $nibble
                        }
	                }";

            var arguments = new object[] { filePathOnTargetMachine, nibbleStart, nibbleSize };

            var bytesRaw = this.RunScriptUsingSession(fetchFileScriptBlock, arguments, runspace, sessionObject);
            var bytes = bytesRaw.Select(_ => (byte)byte.Parse(_.ToString())).ToArray();
            return bytes;
        }

        /// <inheritdoc />
        public string RunCmd(string command, ICollection<string> commandParameters = null)
        {
            var scriptBlock = BuildCmdScriptBlock(command, commandParameters);
            var outputObjects = this.RunScript(scriptBlock);
            var ret = string.Join(Environment.NewLine, outputObjects);
            return ret;
        }

        /// <inheritdoc />
        public string RunCmdOnLocalhost(string command, ICollection<string> commandParameters = null)
        {
            var scriptBlock = BuildCmdScriptBlock(command, commandParameters);
            var outputObjects = this.RunScriptOnLocalhost(scriptBlock);
            var ret = string.Join(Environment.NewLine, outputObjects);
            return ret;
        }

        private static string BuildCmdScriptBlock(string command, ICollection<string> commandParameters)
        {
            var line = " `\"" + command + "`\"";
            foreach (var commandParameter in commandParameters ?? new List<string>())
            {
                line += " `\"" + commandParameter + "`\"";
            }

            line = "\"" + line + "\"";

            var scriptBlock = "{ &cmd.exe /c " + line + " 2>&1 | Write-Output }";
            return scriptBlock;
        }

        /// <inheritdoc />
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times", Justification = "Disposal logic is correct.")]
        public ICollection<dynamic> RunScriptOnLocalhost(string scriptBlock, ICollection<object> scriptBlockParameters = null)
        {
            List<object> ret;

            using (var runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();

                // just send a null session for localhost execution
                ret = this.RunScriptUsingSession(scriptBlock, scriptBlockParameters, runspace, null);

                runspace.Close();
            }

            return ret;
        }

        /// <inheritdoc />
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times", Justification = "Disposal logic is correct.")]
        public ICollection<dynamic> RunScript(string scriptBlock, ICollection<object> scriptBlockParameters = null)
        {
            List<object> ret;

            using (var runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();

                var sessionObject = this.BeginSession(runspace);

                ret = this.RunScriptUsingSession(scriptBlock, scriptBlockParameters, runspace, sessionObject);

                this.EndSession(sessionObject, runspace);

                runspace.Close();
            }

            return ret;
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId = "unneededOutput", Justification = "Prefer to see that output is generated and not used...")]
        private void EndSession(object sessionObject, Runspace runspace)
        {
            if (this.autoManageTrustedHosts)
            {
                RemoveIpAddressFromLocalTrustedHosts(this.IpAddress);
            }

            var removeSessionCommand = new Command("Remove-PSSession");
            removeSessionCommand.Parameters.Add("Session", sessionObject);
            var unneededOutput = RunLocalCommand(runspace, removeSessionCommand);
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "MachineManager", Justification = "Name/spelling is correct.")]
        private object BeginSession(Runspace runspace)
        {
            if (this.autoManageTrustedHosts)
            {
                AddIpAddressToLocalTrustedHosts(this.IpAddress);
            }

            var trustedHosts = GetListOfIpAddressesFromLocalTrustedHosts();
            if (!trustedHosts.Contains(this.IpAddress) && !TrustedHostListIsWildCard(trustedHosts))
            {
                throw new TrustedHostMissingException(
                    "Cannot execute a remote command with out the IP address being added to the trusted hosts list.  Please set MachineManager to handle this automatically or add the address manually: "
                    + this.IpAddress);
            }

            var powershellCredentials = new PSCredential(this.userName, this.password);

            var sessionOptionsCommand = new Command("New-PSSessionOption");
            sessionOptionsCommand.Parameters.Add("OperationTimeout", 0);
            sessionOptionsCommand.Parameters.Add("IdleTimeout", TimeSpan.FromMinutes(20).TotalMilliseconds);
            var sessionOptionsObject = RunLocalCommand(runspace, sessionOptionsCommand).Single().BaseObject;

            var sessionCommand = new Command("New-PSSession");
            sessionCommand.Parameters.Add("ComputerName", this.IpAddress);
            sessionCommand.Parameters.Add("Credential", powershellCredentials);
            sessionCommand.Parameters.Add("SessionOption", sessionOptionsObject);
            var sessionObject = RunLocalCommand(runspace, sessionCommand).Single().BaseObject;
            return sessionObject;
        }

        private List<dynamic> RunScriptUsingSession(
            string scriptBlock,
            ICollection<object> scriptBlockParameters,
            Runspace runspace,
            object sessionObject)
        {
            using (var powershell = PowerShell.Create())
            {
                powershell.Runspace = runspace;

                Collection<PSObject> output;

                // session will implicitly assume remote - if null then localhost...
                if (sessionObject != null)
                {
                    var variableNameArgs = "scriptBlockArgs";
                    var variableNameSession = "invokeCommandSession";
                    powershell.Runspace.SessionStateProxy.SetVariable(variableNameSession, sessionObject);

                    var argsAddIn = string.Empty;
                    if (scriptBlockParameters != null && scriptBlockParameters.Count > 0)
                    {
                        powershell.Runspace.SessionStateProxy.SetVariable(
                            variableNameArgs,
                            scriptBlockParameters.ToArray());
                        argsAddIn = " -ArgumentList $" + variableNameArgs;
                    }

                    var fullScript = "$sc = " + scriptBlock + Environment.NewLine + "Invoke-Command -Session $"
                                     + variableNameSession + argsAddIn + " -ScriptBlock $sc";

                    powershell.AddScript(fullScript);
                    output = powershell.Invoke();
                }
                else
                {
                    var fullScript = "$sc = " + scriptBlock + Environment.NewLine + "Invoke-Command -ScriptBlock $sc";

                    powershell.AddScript(fullScript);
                    foreach (var scriptBlockParameter in scriptBlockParameters ?? new List<object>())
                    {
                        powershell.AddArgument(scriptBlockParameter);
                    }

                    output = powershell.Invoke(scriptBlockParameters);
                }

                this.ThrowOnError(powershell, scriptBlock);

                var ret = output.Cast<dynamic>().ToList();
                return ret;
            }
        }

        private static bool TrustedHostListIsWildCard(IReadOnlyCollection<string> trustedHostList)
        {
            return trustedHostList.Count == 1 && trustedHostList.Single() == "*";
        }

        private static List<PSObject> RunLocalCommand(Runspace runspace, Command arbitraryCommand)
        {
            using (var powershell = PowerShell.Create())
            {
                powershell.Runspace = runspace;

                powershell.Commands.AddCommand(arbitraryCommand);

                var output = powershell.Invoke();

                ThrowOnError(powershell, arbitraryCommand.CommandText, "localhost");

                var ret = output.ToList();
                return ret;
            }
        }

        private void ThrowOnError(PowerShell powershell, string attemptedScriptBlock)
        {
            ThrowOnError(powershell, attemptedScriptBlock, this.IpAddress);
        }

        private static void ThrowOnError(PowerShell powershell, string attemptedScriptBlock, string ipAddress)
        {
            if (powershell.Streams.Error.Count > 0)
            {
                var errorString = string.Join(
                    Environment.NewLine,
                    powershell.Streams.Error.Select(
                        _ =>
                        (_.ErrorDetails == null ? null : _.ErrorDetails.ToString() + " at " + _.ScriptStackTrace)
                        ?? (_.Exception == null ? "Naos.WinRM: No error message available" : _.Exception.ToString() + " at " + _.ScriptStackTrace)));
                throw new RemoteExecutionException(
                    "Failed to run script (" + attemptedScriptBlock + ") on " + ipAddress + " got errors: "
                    + errorString);
            }
        }

        private static string ComputeSha256Hash(byte[] bytes)
        {
            using (var provider = new SHA256Managed())
            {
                var hashBytes = provider.ComputeHash(bytes);
                var calculatedChecksum = string.Empty;

                foreach (byte x in hashBytes)
                {
                    calculatedChecksum += string.Format(CultureInfo.InvariantCulture, "{0:x2}", x);
                }

                return calculatedChecksum;
            }
        }
    }

    /// <summary>
    /// Manages various remote tasks on a machine using the WinRM protocol.
    /// </summary>
    public static class StringExtensions
    {
        /// <summary>
        /// Converts the source string into a secure string. Caller should dispose of the secure string appropriately.
        /// </summary>
        /// <param name="source">The source string.</param>
        /// <returns>A secure version of the source string.</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Caller is expected to dispose of object.")]
        public static SecureString ToSecureString(this string source)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            var result = new SecureString();

            foreach (var character in source.ToCharArray())
            {
                result.AppendChar(character);
            }

            result.MakeReadOnly();

            return result;
        }
    }
}