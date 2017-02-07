using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelToHtmlConverter.ViewModels;
using ExcelToHtmlConverter.Common;
using ExcelToHtmlConverter.Services;
using Moq;
using ExcelToHtmlConverter.Api;

namespace ExcelToHtmlConverterTest
{
    [TestClass]
    public class MainViewModelTest
    {
        private MainViewModel viewModel;
        private Mock<IMessageService> messageMock;
        private Mock<IExcelConverter> converterMock;
        private Mock<IFileService> fileServiceMock;
        private Mock<IDialogService> dialogServiceMock;

        [TestInitialize]
        public void TestInit()
        {
            messageMock = new Mock<IMessageService>();
            converterMock = new Mock<IExcelConverter>();
            fileServiceMock = new Mock<IFileService>();
            dialogServiceMock = new Mock<IDialogService>();
            converterMock.Setup(mock => mock.ConvertWorksheet(It.IsAny<string>()));
            messageMock.Setup(mock => mock.ShowInformation(It.IsAny<string>(), It.IsAny<string>()));
            fileServiceMock.Setup(mock => mock.WriteToFile(It.IsAny<string>(), It.IsAny<string>()));
            DIContainer.Instance.Register<IMessageService>(delegate { return messageMock.Object; });
            DIContainer.Instance.Register<IExcelConverter>(delegate { return converterMock.Object; });
            DIContainer.Instance.Register<IFileService>(delegate { return fileServiceMock.Object; });
            DIContainer.Instance.Register<IDialogService>(delegate { return dialogServiceMock.Object; });
            viewModel = new MainViewModel();
        }

        [TestCleanup]
        public void CleanUp()
        {
            DIContainer.Instance.Reset();
        }

        [TestMethod]
        public void TestThat_CanExecute_Works_FilenameEmpty()
        {
            viewModel.Filename = string.Empty;
            Assert.IsFalse(viewModel.ConvertCommand.CanExecute(null));
            Assert.IsTrue(viewModel.CloseCommand.CanExecute(null));
            Assert.IsTrue(viewModel.InfoCommand.CanExecute(null));
        }

        [TestMethod]
        public void TestThat_CanExecute_Works_FilenameNotEmpty()
        {
            viewModel.Filename = "Filename";
            Assert.IsTrue(viewModel.ConvertCommand.CanExecute(null));
            Assert.IsTrue(viewModel.CloseCommand.CanExecute(null));
            Assert.IsTrue(viewModel.InfoCommand.CanExecute(null));
        }

        [TestMethod]
        public void TestThat_ConvertCommand_CallsLib_AndService()
        {
            viewModel.Filename = "Filename";
            viewModel.Template = null;
            Assert.IsTrue(viewModel.ConvertCommand.CanExecute(null));

            viewModel.ConvertCommand.Execute(null);
            converterMock.Verify(mock => mock.ConvertWorksheet(It.IsAny<string>()), Times.Once());
            fileServiceMock.Verify(mock => mock.WriteToFile(It.Is<string>(a => a == "Filename.html"), It.IsAny<string>()), Times.Once());
        }

        [TestMethod]
        public void TestThat_ConvertCommand_CallsLib_AndService_WithTemplate()
        {
            viewModel.Filename = "Filename";
            viewModel.Template = "Template";
            Assert.IsTrue(viewModel.ConvertCommand.CanExecute(null));

            viewModel.ConvertCommand.Execute(null);
            converterMock.Verify(mock => mock.ConvertWorksheet(It.IsAny<string>(), It.IsAny<string>()), Times.Once());
            fileServiceMock.Verify(mock => mock.WriteToFile(It.Is<string>(a => a == "Filename.html"), It.IsAny<string>()), Times.Once());
        }

        [TestMethod]
        public void TestThat_InfoCommand_CallsService()
        {
            Assert.IsTrue(viewModel.InfoCommand.CanExecute(null));

            viewModel.InfoCommand.Execute(null);
            messageMock.Verify(mock => mock.ShowInformation(It.IsAny<string>(), It.IsAny<string>()), Times.Once());
        }
    }
}
