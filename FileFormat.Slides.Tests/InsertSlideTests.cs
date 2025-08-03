using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.IO;
using System.Collections.Generic;

namespace FileFormat.Slides.Tests
{
    [TestClass]
    public class InsertSlideTests
    {
        protected List<string> tempFileList = new List<string>();

        protected Presentation GetSimplePresentation(int numberOfSlides = 3)
        {
            var tempFile = Path.GetTempFileName();
            var tempPresentationFilename = $"{tempFile}.pptx";

            // track our mess
            tempFileList.Add(tempFile);
            tempFileList.Add(tempPresentationFilename);

            Presentation presentation = Presentation.Create(tempPresentationFilename);
           
            for (var i = 0; i < numberOfSlides; i++)
            {
                var slideId = $"{i + 1}".PadLeft(3, '0');
                Slide slide = new Slide();
                slide.AddTextShapes(new TextShape()
                {
                    Text = $"this is is slide: {slideId}",
                    FontSize = 80

                });
                presentation.AppendSlide(slide);
            }

            return presentation;
        }

        protected Slide BuildSimpleSlide()
        {
            Slide slide = new Slide();
            slide.AddTextShapes(new TextShape()
            {
                Text = "this is is slide: 009",
                FontSize = 80

            });
            return slide;
        }

        protected string[] GetRelationshipIdsInOrder(Presentation presentation) 
        { 
            return presentation.Facade.SlideIdList.ChildElements
                        .Select(s => (s as DocumentFormat.OpenXml.Presentation.SlideId).RelationshipId.Value)
                        .ToArray();
        }

        [TestCleanup]
        public void CleanUp() 
        {
            if (tempFileList.Any())
            {
                foreach(var file in tempFileList)
                {
                    try {
                        File.Delete(file);
                    }
                    catch {
                        Console.WriteLine($"Failed to delete: {file}");
                    }                    
                }
            }
        }

        [TestMethod]
        public void InsertSlideAtPositionZero()
        {
            var presentation = GetSimplePresentation();
            var idList = GetRelationshipIdsInOrder(presentation);

            // insert a new slide at the begining of the presentation
            Slide slide = BuildSimpleSlide();
            presentation.InsertSlideAt(0, slide);

            // get the new relationship id from it to build out the expected list
            var newSlideRelationshipId = slide.RelationshipId;
            var expectedIds = new string[] { newSlideRelationshipId, idList[0], idList[1], idList[2] };

            // Get the current list post insert
            var updatedIdList = GetRelationshipIdsInOrder(presentation);

            // validate the slide is where we expext it to be
            CollectionAssert.AreEqual(expectedIds, updatedIdList);

            presentation.Save();
            presentation.close();
        }

        [TestMethod]
        public void InsertSlideAtPositionOne()
        {
            var presentation = GetSimplePresentation();
            var idList = GetRelationshipIdsInOrder(presentation);

            // insert a new slide at the begining of the presentation
            Slide slide = BuildSimpleSlide();
            presentation.InsertSlideAt(1, slide);

            // get the new relationship id from it to build out the expected list
            var newSlideRelationshipId = slide.RelationshipId;
            var expectedIds = new string[] { idList[0], newSlideRelationshipId, idList[1], idList[2] };

            // Get the current list post insert
            var updatedIdList = GetRelationshipIdsInOrder(presentation);

            // validate the slide is where we expext it to be
            CollectionAssert.AreEqual(expectedIds, updatedIdList);

            presentation.Save();
            presentation.close();
        }

        [TestMethod]
        public void InsertSlideAtPositionTwo()
        {
            var presentation = GetSimplePresentation();
            var idList = GetRelationshipIdsInOrder(presentation);

            // insert a new slide at the begining of the presentation
            Slide slide = BuildSimpleSlide();
            presentation.InsertSlideAt(2, slide);

            // get the new relationship id from it to build out the expected list
            var newSlideRelationshipId = slide.RelationshipId;
            var expectedIds = new string[] { idList[0], idList[1], newSlideRelationshipId, idList[2] };

            // Get the current list post insert
            var updatedIdList = GetRelationshipIdsInOrder(presentation);

            // validate the slide is where we expext it to be
            CollectionAssert.AreEqual(expectedIds, updatedIdList);

            presentation.Save();
            presentation.close();
        }

        [TestMethod]
        public void InsertSlideAtPositionThree()
        {
            var presentation = GetSimplePresentation();
            var idList = GetRelationshipIdsInOrder(presentation);

            // insert a new slide at the begining of the presentation
            Slide slide = BuildSimpleSlide();
            presentation.InsertSlideAt(3, slide);

            // get the new relationship id from it to build out the expected list
            var newSlideRelationshipId = slide.RelationshipId;
            var expectedIds = new string[] { idList[0], idList[1], idList[2], newSlideRelationshipId };

            // Get the current list post insert
            var updatedIdList = GetRelationshipIdsInOrder(presentation);

            // validate the slide is where we expext it to be
            CollectionAssert.AreEqual(expectedIds, updatedIdList);

            presentation.Save();
            presentation.close();
        }

        [TestMethod]
        public void InsertSlideAtPositionOutOfRangeHigh()
        {
            var presentation = GetSimplePresentation();
            var idList = GetRelationshipIdsInOrder(presentation);

            // insert a new slide at the begining of the presentation
            Slide slide = BuildSimpleSlide();
            presentation.InsertSlideAt(99, slide);

            // get the new relationship id from it to build out the expected list
            var newSlideRelationshipId = slide.RelationshipId;
            var expectedIds = new string[] { idList[0], idList[1], idList[2], newSlideRelationshipId };

            // Get the current list post insert
            var updatedIdList = GetRelationshipIdsInOrder(presentation);

            // validate the slide is where we expext it to be
            CollectionAssert.AreEqual(expectedIds, updatedIdList);

            presentation.Save();
            presentation.close();
        }

        [TestMethod]
        public void InsertSlideAtPositionOutOfRangeLow()
        {
            var presentation = GetSimplePresentation();
            var idList = GetRelationshipIdsInOrder(presentation);

            // insert a new slide at the begining of the presentation
            Slide slide = BuildSimpleSlide();
            presentation.InsertSlideAt(-5, slide);

            // get the new relationship id from it to build out the expected list
            var newSlideRelationshipId = slide.RelationshipId;
            var expectedIds = new string[] { idList[0], idList[1], idList[2], newSlideRelationshipId };

            // Get the current list post insert
            var updatedIdList = GetRelationshipIdsInOrder(presentation);

            // validate the slide is where we expext it to be
            CollectionAssert.AreEqual(expectedIds, updatedIdList);

            presentation.Save();
            presentation.close();
        }

    }
}
