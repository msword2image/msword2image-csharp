namespace MsWordToImage
{
    public class Input
    {
        private InputType type;
        private string value;

        public Input(InputType type, string value)
        {
            this.type = type;
            this.value = value;
        }

        public InputType getType()
        {
            return this.type;
        }

        public string getValue()
        {
            return this.value;
        }
    }
}
