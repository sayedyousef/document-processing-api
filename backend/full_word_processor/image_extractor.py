# backend/processors/full-word-processor/image_extractor.py

from pathlib import Path
from .models import ProcessingContext, ImageInfo

class ImageExtractor:
    """Handles image extraction and setup"""
    
    def setup(self, context: ProcessingContext) -> ProcessingContext:
        """Setup image extraction directory"""
        
        print("\nSTEP 4: Setting up image extraction...")
        
        # Create images directory
        context.images_dir = context.output_dir / f"{context.input_path.stem}_images"
        context.images_dir.mkdir(exist_ok=True)
        
        print(f"✓ Images folder: {context.images_dir}")
        
        return context
    
    def extract_image(self, image_data, context: ProcessingContext) -> str:
        """Extract single image and return relative path"""
        
        try:
            # Generate filename
            image_num = len(list(context.images_dir.glob('*'))) + 1
            extension = image_data.content_type.split('/')[-1]
            filename = f"image_{image_num}.{extension}"
            
            # Save image
            image_path = context.images_dir / filename
            with open(image_path, 'wb') as f:
                f.write(image_data.read())
            
            # Store info
            context.images.append(ImageInfo(
                filename=filename,
                path=image_path,
                content_type=image_data.content_type
            ))
            
            # Return relative path
            return f"{context.images_dir.name}/{filename}"
            
        except Exception as e:
            print(f"⚠ Image extraction failed: {e}")
            return ""