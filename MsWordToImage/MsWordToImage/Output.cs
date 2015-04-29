namespace MsWordToImage
{
    public class Output
    {
        private OutputType type;
        private OutputImageFormat imageFormat;
        private string value;

        public Output(OutputType type, OutputImageFormat imageFormat)
            : this(type, imageFormat, null)
        {
        }

        public Output(OutputType type, OutputImageFormat imageFormat, string value)
        {
            this.type = type;
            this.imageFormat = imageFormat;
            this.value = value;
        }

        public OutputType getType()
        {
            return this.type;
        }

        public OutputImageFormat getImageFormat()
        {
            return this.imageFormat;
        }

        public string getValue()
        {
            return this.value;
        }
    }
}
