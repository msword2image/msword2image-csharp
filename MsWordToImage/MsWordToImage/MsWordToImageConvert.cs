namespace MsWordToImage
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Net.Http;
    using System.Web;

    /// <summary>
    /// This class converts microsoft word documents to image
    /// </summary>
    public class MsWordToImageConvert
    {
        private string apiUser;
        private string apiKey;
        private Output output;
        private Input input;

        /// <summary>
        /// The constructors for conversion
        /// </summary>
        /// <param name="apiUser">The API username given by msword2image.com</param>
        /// <param name="apiKey">The API password given by msword2image.com</param>
        public MsWordToImageConvert(string apiUser, string apiKey)
        {
            this.apiUser = apiUser;
            this.apiKey = apiKey;
            this.input = null;
            this.output = null;
        }

        /// <summary>
        /// Convert from file with given name
        /// </summary>
        /// <param name="filename">The filename given</param>
        public void fromFile(string filename)
        {
            this.input = new Input(InputType.File, filename);
        }

        /// <summary>
        /// Convert from given URL 
        /// </summary>
        /// <param name="url">The URL of the word document</param>
        public void fromURL(string url)
        {
            this.input = new Input(InputType.URL, url);
        }

        /// <summary>
        /// Convert to file with given name in JPEG format
        /// </summary>
        /// <param name="filename">The output file name to put the result image</param>
        /// <returns>True on success, false on error</returns>
        public bool toFile(string filename)
        {
            return this.toFile(filename, OutputImageFormat.JPEG);
        }

        /// <summary>
        /// Converts to file with given name and format
        /// </summary>
        /// <param name="filename">The output file name to save the image to</param>
        /// <param name="imageFormat">The output image file format. Example: JPEG, GIF, PNG</param>
        /// <returns>True on success, false on error</returns>
        public bool toFile(string filename, OutputImageFormat imageFormat)
        {
            this.output = new Output(OutputType.File, imageFormat, filename);
            return this.convertToFile();
        }

        /// <summary>
        /// Converts to base64 encoded string that represents the JPEG image
        /// </summary>
        /// <returns>The base64 encoded image</returns>
        public string toBase46EncodedString()
        {
            return this.toBase46EncodedString(OutputImageFormat.JPEG);
        }

        /// <summary>
        /// Converts to base64 encoded string with given output image format
        /// </summary>
        /// <param name="imageFormat">The desired output image format. Example: JPEG, GIF, PNG</param>
        /// <returns>The base64 encoded string that represents the image</returns>
        public string toBase46EncodedString(OutputImageFormat imageFormat)
        {
            this.output = new Output(OutputType.Base64EncodedString, imageFormat);
            return this.convertToBase46EncodedString();
        }

        /// <summary>
        /// Checks if the input was set
        /// </summary>
        private void checkInput()
        {
            if (this.input == null)
            {
                throw new ArgumentException("MsWordToImageConvert: Input was not set!");
            }
        }

        /// <summary>
        /// Checks if the output was set
        /// </summary>
        private void checkOutput()
        {
            if (this.output == null)
            {
                throw new ArgumentException("MsWordToImageConvert: Output was not set!");
            }
        }

        /// <summary>
        /// Gets all allowed input/output combinations
        /// </summary>
        /// <returns>allowed input/output combinations</returns>
        private Dictionary<InputType, HashSet<OutputType>> getAllowedInputOutputTypes()
        {
            Dictionary<InputType, HashSet<OutputType>> returnValue = new Dictionary<InputType, HashSet<OutputType>>()
            {
                {
                    InputType.URL,
                    new HashSet<OutputType>() {
                        OutputType.File,
                        OutputType.Base64EncodedString
                    }
                },
                {
                    InputType.File,
                    new HashSet<OutputType>() {
                        OutputType.File,
                        OutputType.Base64EncodedString
                    }
                },
            };
            return returnValue;
        }

        /// <summary>
        /// Checks if provided input/output conversion is supported
        /// </summary>
        private void checkInputAndOutputTypes()
        {
            if (!this.getAllowedInputOutputTypes().ContainsKey(this.input.getType()))
            {
                throw new ArgumentException("MsWordToImageConvert: Conversion from " + this.input.getType().ToString() + " is not supported");
            }

            HashSet<OutputType> supportedOutputTypes = this.getAllowedInputOutputTypes()[this.input.getType()];
            if (!supportedOutputTypes.Contains(this.output.getType()))
            {
                throw new ArgumentException("MsWordToImageConvert: Conversion from " + this.input.getType().ToString() + " to " + this.output.getType().ToString() + " is not supported");
            }
        }

        /// <summary>
        /// Checks if both input and output is set
        /// </summary>
        private void checkInputAndOutput()
        {
            this.checkInput();
            this.checkOutput();
        }

        /// <summary>
        /// Checks if both input and output is set
        /// And also checks if given input/output combination is valid
        /// </summary>
        private void checkConversionSanity()
        {
            this.checkInputAndOutput();
            this.checkInputAndOutputTypes();
        }

        /// <summary>
        /// Constructs the msword2image.com address with three parameters. These are:
        /// 1. apiUser: given by msword2image.com
        /// 2. apiKey: given by msword2image.com
        /// 3. format: the desired output image format
        /// </summary>
        /// <param name="additionalParameters">The additional query string parameters</param>
        /// <returns>The constructed URL that points to msword2image.com</returns>
        private String constructMsWordToImageAddress(Dictionary<String, String> additionalParameters)
        {
            String returnValue = "http://msword2image.com/convert?"
                + "apiUser=" + HttpUtility.UrlEncode(this.apiUser) + "&"
                + "apiKey=" + HttpUtility.UrlEncode(this.apiKey) + "&"
                + "format=" + HttpUtility.UrlEncode(this.output.getImageFormat().ToString());

            foreach (KeyValuePair<string, string> entry in additionalParameters)
            {
                returnValue += "&" + entry.Key + "=" + HttpUtility.UrlEncode(entry.Value);
            }

            return returnValue;
        }

        /// <summary>
        /// Converts from * to file
        /// </summary>
        /// <returns>True on success, false on error.</returns>
        private bool convertToFile()
        {
            this.checkConversionSanity();

            if (this.input.getType().Equals(InputType.File))
            {
                return this.convertFromFileToFile();
            }
            else if (this.input.getType().Equals(InputType.URL))
            {
                return this.convertFromURLToFile();
            }

            return false;
        }

        /// <summary>
        /// Gets output as <see cref="FileInfo"/> object
        /// </summary>
        /// <returns>The constructed <see cref="FileInfo"/> object</returns>
        private FileInfo getOutputFile()
        {
            return new FileInfo(this.output.getValue());
        }

        /// <summary>
        /// Converts from URL to output file
        /// </summary>
        /// <returns>True on success, false on error.</returns>
        private bool convertFromURLToFile()
        {
            return this.convertFromURLToFile(this.getOutputFile());
        }

        /// <summary>
        /// Converts from URL to given output file destination
        /// </summary>
        /// <param name="dest">The output destination to save the image to</param>
        /// <returns>True on success, false on error.</returns>
        private bool convertFromURLToFile(FileInfo dest)
        {
            Dictionary<String, String> parameters = new Dictionary<String, String>()
            {
                { "url", this.input.getValue() }
            };

            using (WebClient client = new WebClient())
            {
                client.DownloadFile(this.constructMsWordToImageAddress(parameters), dest.FullName);
            }
            return true;
        }

        /// <summary>
        /// Converts from * to base64 encoded image string
        /// </summary>
        /// <returns>The base64 encoded image string</returns>
        private string convertToBase46EncodedString()
        {
            this.checkConversionSanity();

            if (this.input.getType().Equals(InputType.File))
            {
                return this.convertFromFileToBase46EncodedString();
            }
            else if (this.input.getType().Equals(InputType.URL))
            {
                return this.convertFromURLToBase46EncodedString();
            }
            else
            {
                throw new ArgumentException("MsWordToImageConvert: Conversion from " + this.input.getType().ToString() + " is not supported");
            }
        }

        /// <summary>
        /// Creates a temp file and returns pointer <see cref="FileInfo"/> to it.
        /// </summary>
        /// <returns>The <see cref="FileInfo"/> object that represents the temp file</returns>
        private FileInfo getTempFile()
        {
            string tempFileName = Path.GetTempFileName();
            return new FileInfo(tempFileName);
        }

        /// <summary>
        /// Base64 encodes a given file and returns it as string
        /// </summary>
        /// <param name="source">The source file to base64 encode</param>
        /// <returns>The base64 encoded file contents</returns>
        private string base64EncodeContentsOfFile(FileInfo source)
        {
            byte[] outputByteArray = File.ReadAllBytes(source.FullName);
            return System.Convert.ToBase64String(outputByteArray);
        }

        /// <summary>
        /// Converts from URL to base64 encoded string
        /// </summary>
        /// <returns>The base64 encoded string that represents the output image</returns>
        private string convertFromURLToBase46EncodedString()
        {
            FileInfo tempFile = this.getTempFile();
            this.convertFromURLToFile(tempFile);
            return this.base64EncodeContentsOfFile(tempFile);
        }

        /// <summary>
        /// Gets the input file
        /// </summary>
        /// <returns>The <see cref="FileInfo"/> that represents the input file object</returns>
        private FileInfo getInputFile()
        {
            return new FileInfo(this.input.getValue());
        }

        /// <summary>
        /// Actual business logic. Converts from File word document to file image.
        /// </summary>
        /// <param name="dest">The destination file object</param>
        /// <returns>True on success, false on error.</returns>
        private bool convertFromFileToFile(FileInfo dest)
        {
            if (!this.getInputFile().Exists)
            {
                throw new ArgumentException("MsWordToImageConvert: Input file was not found at '" + this.input.getValue() + "'");
            }

            using (var client = new HttpClient())
            {
                using (var formData = new MultipartFormDataContent())
                {
                    HttpContent fileStreamContent = new StreamContent(new FileStream(this.getInputFile().FullName, FileMode.Open));
                    formData.Add(fileStreamContent, "file_contents");

                    var response = client.PostAsync(this.constructMsWordToImageAddress(new Dictionary<string, string>()), formData).Result;
                    if (!response.IsSuccessStatusCode)
                    {
                        return false;
                    }

                    Stream resultStream = response.Content.ReadAsStreamAsync().Result;
                    using (var fileStream = File.Create(dest.FullName))
                    {
                        resultStream.Seek(0, SeekOrigin.Begin);
                        resultStream.CopyTo(fileStream);
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Converts from input word document to output image
        /// </summary>
        /// <returns>True on success, false on error</returns>
        private bool convertFromFileToFile()
        {
            return this.convertFromFileToFile(this.getOutputFile());
        }

        /// <summary>
        /// Converts from input word file to base64 encoded string
        /// </summary>
        /// <returns>The base64 encoded image that represents the file</returns>
        private string convertFromFileToBase46EncodedString()
        {
            FileInfo tempFile = this.getTempFile();
            this.convertFromFileToFile(tempFile);
            return this.base64EncodeContentsOfFile(tempFile);
        }
    }
}
