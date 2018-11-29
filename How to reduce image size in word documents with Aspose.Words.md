# How to reduce image size in word documents with `` Aspose.Words ``

```c#
public void Reduce(string file, long value, Guid imageFormat = ImageFormat.Jpeg.Guid){
    Document doc = new Document(file);
    NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
    
    var myEncoderParameters = new EncoderParameters();
    var encoder = System.Drawing.Imaging.Encoder.Quality;
    var myEncoderParameter = new EncoderParameter(encoder, value);
    myEncoderParameters.Param[0] = myEncoderParameter;
    
    ImageCodecInfo myImageCodecInfo = ImageCodecInfo.GetImageEncoders().Where(q => q.FormatID == imageFormat).FirstOrDefault();
    
    int index = 0;
    foreach(Shape shape in shapes){
        if(shape.HasImage){
            Image image = shape.ImageData.ToImage();
            string tmpImage = $"tmpImage_{index++}.jpeg";
            image.Save(tmpImage, myImageCodecInfo, myEncoderParameters);
            FileInfo fi = new FileInfo(tmpImage);
            shape.ImageData.SetImage(fi.FullName);
            fi.Delete();
        }
    }
    
    doc.Save($"{Path.GetFileNameWithoutExtension}_strip.{Path.GetExtension(file)}");
}
```

