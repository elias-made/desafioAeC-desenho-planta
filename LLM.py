import os
from dotenv import load_dotenv
from pydantic_ai.models.bedrock import BedrockConverseModel
from pydantic_ai.providers.bedrock import BedrockProvider
import boto3
from botocore.config import Config

load_dotenv()

def get_model():
    client = boto3.client(
        'bedrock-runtime',
        region_name=os.getenv('AWS_DEFAULT_REGION', 'us-east-1'),
        aws_access_key_id=os.getenv('AWS_ACCESS_KEY_ID'),
        aws_secret_access_key=os.getenv('AWS_SECRET_ACCESS_KEY'),
        config=Config(
            read_timeout=300,      # 5 minutos para respostas longas
            connect_timeout=10,
            retries={'max_attempts': 2},
        ),
    )
    return BedrockConverseModel(
        'us.anthropic.claude-sonnet-4-5-20250929-v1:0',
        provider=BedrockProvider(bedrock_client=client),
    )

planner_model = get_model()
